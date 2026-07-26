//! Flow file I/O (desktop).
//!
//! A flow is a `.ebb` file the user owns, opens, and backs up with the rest of
//! their filesystem. This module is the byte layer for those files: it resolves
//! the default flows directory, reads and writes flow text, keeps the
//! recent-flows list beside the config file, and buffers paths the OS asks us to
//! open. It knows nothing about the format - the frontend parses and validates.
//!
//! The filesystem plugin is deliberately not used. Its scope model cannot
//! express "whatever path the user picked, remembered across restarts" without
//! persisting an ever-growing path allowlist to disk, which is a poor trade
//! against granting the webview a general filesystem API. These five commands
//! are the whole surface instead, and the recents path is fixed here rather than
//! passed in, so it cannot be steered from JS.

use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use parking_lot::Mutex;
use serde::Serialize;
use tauri::{AppHandle, Emitter, Manager, State};

const EXT: &str = "ebb";

/// Where the start screen files new flows, and the home dir it shortens for
/// display. Resolution only; nothing is created until a flow is written.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FlowPaths {
    flows_dir: String,
    home: String,
}

fn home_dir() -> Option<PathBuf> {
    #[cfg(windows)]
    {
        std::env::var_os("USERPROFILE").map(PathBuf::from)
    }
    #[cfg(not(windows))]
    {
        std::env::var_os("HOME").map(PathBuf::from)
    }
}

#[tauri::command]
pub fn flow_paths() -> Result<FlowPaths, String> {
    let home = home_dir().ok_or("Could not locate your home directory")?;
    Ok(FlowPaths {
        flows_dir: home
            .join("Documents")
            .join("ebb")
            .to_string_lossy()
            .into_owned(),
        home: home.to_string_lossy().into_owned(),
    })
}

// --- Flow files ----------------------------------------------------------------

/// Write via a temp file in the same directory, then rename over the target.
/// The rename is the atomic step, but only if the bytes are already durable, so
/// the temp file is synced before it is moved: a crash mid-round can cost the
/// last debounce interval, never a truncated flow.
fn write_atomic(path: &Path, contents: &str) -> Result<(), String> {
    let dir = path
        .parent()
        .ok_or_else(|| format!("{} has no parent directory", path.display()))?;
    fs::create_dir_all(dir).map_err(|e| format!("Could not create {}: {e}", dir.display()))?;

    let name = path
        .file_name()
        .and_then(|n| n.to_str())
        .ok_or_else(|| format!("{} has no filename", path.display()))?;
    let tmp = dir.join(format!(".{name}.tmp"));

    let write = || -> std::io::Result<()> {
        let mut f = fs::File::create(&tmp)?;
        f.write_all(contents.as_bytes())?;
        f.sync_all()
    };
    if let Err(e) = write() {
        let _ = fs::remove_file(&tmp);
        return Err(format!("Could not save {}: {e}", path.display()));
    }

    fs::rename(&tmp, path).map_err(|e| {
        let _ = fs::remove_file(&tmp);
        format!("Could not save {}: {e}", path.display())
    })
}

/// `None` means the file is gone, which is an ordinary outcome: a recent entry
/// whose flow was moved or deleted is dropped from the list rather than raised
/// as an error.
#[tauri::command]
pub fn read_flow_file(path: String) -> Result<Option<String>, String> {
    match fs::read_to_string(&path) {
        Ok(text) => Ok(Some(text)),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(e) => Err(format!("Could not read {path}: {e}")),
    }
}

#[tauri::command]
pub fn write_flow_file(path: String, contents: String) -> Result<(), String> {
    write_atomic(Path::new(&path), &contents)
}

/// Create a flow without ever overwriting one, returning the path actually
/// used. Deduping here rather than in JS keeps the check and the create in one
/// step, so two flows made in the same second cannot race onto one filename.
#[tauri::command]
pub fn create_flow_file(dir: String, name: String, contents: String) -> Result<String, String> {
    let dir = PathBuf::from(dir);
    fs::create_dir_all(&dir).map_err(|e| format!("Could not create {}: {e}", dir.display()))?;

    let stem = name
        .strip_suffix(&format!(".{EXT}"))
        .unwrap_or(&name)
        .to_string();

    for n in 1..1000 {
        let candidate = if n == 1 {
            format!("{stem}.{EXT}")
        } else {
            format!("{stem}-{n}.{EXT}")
        };
        let path = dir.join(&candidate);
        match fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)
        {
            Ok(mut f) => {
                f.write_all(contents.as_bytes())
                    .and_then(|()| f.sync_all())
                    .map_err(|e| format!("Could not write {}: {e}", path.display()))?;
                return Ok(path.to_string_lossy().into_owned());
            }
            Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(e) => return Err(format!("Could not create {}: {e}", path.display())),
        }
    }
    Err(format!("Too many flows already named {stem}"))
}

// --- Recent flows ----------------------------------------------------------------

fn recents_path() -> Option<PathBuf> {
    crate::config::config_dir().map(|d| d.join("recents.json"))
}

#[tauri::command]
pub fn read_recents() -> Result<Option<String>, String> {
    let Some(path) = recents_path() else {
        return Ok(None);
    };
    match fs::read_to_string(&path) {
        Ok(text) => Ok(Some(text)),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(e) => Err(format!("Could not read {}: {e}", path.display())),
    }
}

#[tauri::command]
pub fn write_recents(contents: String) -> Result<(), String> {
    let path = recents_path().ok_or("Could not locate your config directory")?;
    write_atomic(&path, &contents)
}

// --- Open-with ---------------------------------------------------------------------

#[derive(Default)]
struct PendingState {
    queue: Vec<String>,
    drained: bool,
}

/// Paths the OS handed us before the frontend was listening. macOS delivers
/// these as `RunEvent::Opened`, Windows and Linux as argv; either way they can
/// arrive before the webview exists, so they wait here until it drains them.
#[derive(Default)]
pub struct PendingOpen(Mutex<PendingState>);

/// Route an OS-requested path to the frontend: buffered before the first drain,
/// emitted live after it. Both sides of the decision are made under one lock so
/// a path arriving mid-drain cannot fall between them and be lost.
pub fn request_open(app: &AppHandle, path: String) {
    let state = app.state::<PendingOpen>();
    let mut pending = state.0.lock();
    if pending.drained {
        drop(pending);
        let _ = app.emit("file:open", path);
    } else {
        pending.queue.push(path);
    }
}

#[tauri::command]
pub fn drain_pending_open(state: State<'_, PendingOpen>) -> Vec<String> {
    let mut pending = state.0.lock();
    pending.drained = true;
    std::mem::take(&mut pending.queue)
}

/// The openable file paths in a command line, ignoring the executable and any
/// flags. Used for the launch argv and again for a second instance's argv.
pub fn flow_paths_in(args: &[String]) -> Vec<String> {
    args.iter()
        .skip(1)
        .filter(|a| !a.starts_with('-'))
        .filter(|a| Path::new(a).is_file())
        .cloned()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The frontend reads `flowsDir`. Serde defaults to the field name, so
    /// without the camelCase rename this struct silently hands over
    /// `flows_dir`, the frontend reads undefined, and every new flow is
    /// created with no directory. Nothing in the type system catches that -
    /// only this does.
    #[test]
    fn paths_cross_the_wire_in_camel_case() {
        let json = serde_json::to_value(FlowPaths {
            flows_dir: "/home/a/Documents/ebb".into(),
            home: "/home/a".into(),
        })
        .unwrap();

        assert_eq!(json["flowsDir"], "/home/a/Documents/ebb");
        assert_eq!(json["home"], "/home/a");
        assert!(json.get("flows_dir").is_none());
    }

    fn tmpdir(tag: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("ebb-flowfile-{tag}-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn atomic_write_leaves_no_temp_file() {
        let dir = tmpdir("atomic");
        let path = dir.join("round.ebb");
        write_atomic(&path, "{\"version\":3}").unwrap();

        assert_eq!(fs::read_to_string(&path).unwrap(), "{\"version\":3}");
        let leftovers: Vec<_> = fs::read_dir(&dir)
            .unwrap()
            .map(|e| e.unwrap().file_name().to_string_lossy().into_owned())
            .filter(|n| n.ends_with(".tmp"))
            .collect();
        assert!(leftovers.is_empty(), "left behind {leftovers:?}");
    }

    #[test]
    fn atomic_write_replaces_without_truncating() {
        let dir = tmpdir("replace");
        let path = dir.join("round.ebb");
        write_atomic(&path, "first").unwrap();
        write_atomic(&path, "second").unwrap();
        assert_eq!(fs::read_to_string(&path).unwrap(), "second");
    }

    #[test]
    fn create_never_overwrites_an_existing_flow() {
        let dir = tmpdir("create");
        let d = dir.to_string_lossy().into_owned();

        let first = create_flow_file(d.clone(), "policy-2026-07-25.ebb".into(), "a".into()).unwrap();
        let second =
            create_flow_file(d.clone(), "policy-2026-07-25.ebb".into(), "b".into()).unwrap();

        assert!(first.ends_with("policy-2026-07-25.ebb"));
        assert!(second.ends_with("policy-2026-07-25-2.ebb"));
        assert_eq!(fs::read_to_string(&first).unwrap(), "a");
        assert_eq!(fs::read_to_string(&second).unwrap(), "b");
    }

    #[test]
    fn reading_a_missing_flow_is_not_an_error() {
        let dir = tmpdir("missing");
        let path = dir.join("gone.ebb").to_string_lossy().into_owned();
        assert_eq!(read_flow_file(path).unwrap(), None);
    }

    #[test]
    fn argv_keeps_only_real_files() {
        let dir = tmpdir("argv");
        let flow = dir.join("round.ebb");
        fs::write(&flow, "{}").unwrap();
        let flow = flow.to_string_lossy().into_owned();

        let args = vec![
            "/Applications/ebb.app/Contents/MacOS/ebb".to_string(),
            "--flag".to_string(),
            dir.join("absent.ebb").to_string_lossy().into_owned(),
            flow.clone(),
        ];
        assert_eq!(flow_paths_in(&args), vec![flow]);
    }
}
