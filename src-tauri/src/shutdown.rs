//! Flush before a window closes or the app exits.
//!
//! Autosave debounces, so at any instant the newest edit may still be in
//! memory. Closing a window or quitting used to tear the process down
//! immediately, taking that edit with it - the one outcome the product exists
//! to prevent. Every close now runs through here instead: the frontend is
//! asked to write, and only then does the window actually close (or, for a
//! full quit, the process actually exit).
//!
//! Two shapes of attempt share one piece of state:
//!   - **Closing one window** claims just that window's label. Once it
//!     confirms (or times out), that window alone is destroyed.
//!   - **Quitting** (Cmd+Q, or closing the last open window) claims every
//!     open window's label at once. Exit only happens once every one of them
//!     has confirmed; one window reporting a failed write cancels the whole
//!     attempt; a lone timeout deadline covers the group.
//!
//! Either way, a claimed label is never released on success - there is
//! nothing left to release it for, since the window is about to be destroyed
//! or the whole process is about to exit, and the OS can synthesize another
//! close request for a window mid-teardown. Only a reported failure releases
//! a claim, because that is the one outcome that leaves the window(s) open
//! and needing to be closable again.

use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::LazyLock;
use std::time::Duration;

use parking_lot::Mutex;
use tauri::{AppHandle, Emitter, Manager, Runtime, WebviewWindow};

use crate::windows;

/// How long a window stays open waiting for its frontend to confirm.
const FLUSH_TIMEOUT: Duration = Duration::from_secs(3);

static NEXT_ATTEMPT: AtomicU64 = AtomicU64::new(0);

/// What happens once every window in an attempt has confirmed.
enum Action {
    ExitProcess,
    CloseWindow(String),
}

struct Attempt {
    /// Every label this attempt covers, fixed for its lifetime - a
    /// cancellation needs this to release windows that already confirmed,
    /// not just the ones still waiting.
    labels: HashSet<String>,
    /// Labels not yet confirmed.
    pending: HashSet<String>,
    action: Action,
}

#[derive(Default)]
struct State {
    /// The attempt currently claiming each window label. Entries are removed
    /// only by a cancellation; a successful or timed-out attempt leaves them
    /// in place, since the window it covered is gone or the process is
    /// exiting either way.
    owner: HashMap<String, u64>,
    attempts: HashMap<u64, Attempt>,
}

static STATE: LazyLock<Mutex<State>> = LazyLock::new(|| Mutex::new(State::default()));

/// Claims every label in `labels` under one new attempt, or claims nothing
/// and returns `None` if any of them is already mid-flush - hammering a
/// close button, or Cmd+Q while a window-close is in flight, must not stack
/// a second timeout over the same window.
fn claim(labels: &[String], action: Action) -> Option<u64> {
    let mut state = STATE.lock();
    if labels.iter().any(|l| state.owner.contains_key(l)) {
        return None;
    }
    let attempt = NEXT_ATTEMPT.fetch_add(1, Ordering::SeqCst);
    let set: HashSet<String> = labels.iter().cloned().collect();
    for l in labels {
        state.owner.insert(l.clone(), attempt);
    }
    state.attempts.insert(
        attempt,
        Attempt {
            labels: set.clone(),
            pending: set,
            action,
        },
    );
    Some(attempt)
}

/// Marks `label` done under its current attempt; returns the attempt's
/// action once every label it covers has confirmed, `None` otherwise (either
/// more are still pending, or the attempt already resolved).
fn resolve_one(label: &str) -> Option<Action> {
    let mut state = STATE.lock();
    let attempt = *state.owner.get(label)?;
    let entry = state.attempts.get_mut(&attempt)?;
    entry.pending.remove(label);
    if entry.pending.is_empty() {
        state.attempts.remove(&attempt).map(|a| a.action)
    } else {
        None
    }
}

/// Overrides a hung reply for an entire attempt: whatever is still pending
/// completes as if every remaining window had confirmed.
fn resolve_timeout(attempt: u64) -> Option<Action> {
    STATE.lock().attempts.remove(&attempt).map(|a| a.action)
}

/// Abandons the attempt `label` belongs to, releasing every label it covers
/// (including ones that already confirmed) so they can be closed or quit
/// again later. The one outcome a failed write must produce: nothing closes,
/// nothing exits.
fn cancel(label: &str) {
    let mut state = STATE.lock();
    let Some(&attempt) = state.owner.get(label) else {
        return;
    };
    if let Some(entry) = state.attempts.remove(&attempt) {
        for l in &entry.labels {
            state.owner.remove(l);
        }
    }
}

/// True while `label` is mid-flush, under either kind of attempt.
pub fn in_progress(label: &str) -> bool {
    STATE.lock().owner.contains_key(label)
}

fn arm_timeout<R: Runtime>(app: &AppHandle<R>, attempt: u64) {
    let handle = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(FLUSH_TIMEOUT);
        if let Some(action) = resolve_timeout(attempt) {
            complete(&handle, action);
        }
    });
}

fn complete<R: Runtime>(app: &AppHandle<R>, action: Action) {
    match action {
        Action::ExitProcess => app.exit(0),
        Action::CloseWindow(label) => {
            windows::note_close(&label);
            if let Some(w) = app.get_webview_window(&label) {
                let _ = w.destroy();
            }
        }
    }
}

/// Begins closing exactly one window: asks its frontend to flush, then
/// destroys it once confirmed (never exits the process).
pub fn request_window<R: Runtime>(app: &AppHandle<R>, window: &WebviewWindow<R>) {
    let label = window.label().to_string();
    let Some(attempt) = claim(
        std::slice::from_ref(&label),
        Action::CloseWindow(label.clone()),
    ) else {
        return;
    };
    let _ = app.emit_to(window.label(), "app:flush", ());
    arm_timeout(app, attempt);
}

/// Begins a full quit: asks every open window's frontend to flush, and exits
/// only once all of them confirm.
pub fn request_all<R: Runtime>(app: &AppHandle<R>) {
    let labels: Vec<String> = app.webview_windows().keys().cloned().collect();
    if labels.is_empty() {
        app.exit(0);
        return;
    }
    let Some(attempt) = claim(&labels, Action::ExitProcess) else {
        return;
    };
    // Every window must see this, so it stays a broadcast rather than a
    // single `target_window` emit.
    let _ = app.emit("app:flush", ());
    arm_timeout(app, attempt);
}

/// The frontend's answer, from the window that sent it. `saved` false keeps
/// that window (and, in a full quit, every other window in the same attempt)
/// open with the round intact; the frontend is responsible for saying why.
#[tauri::command]
pub fn finish_quit<R: Runtime>(app: AppHandle<R>, window: WebviewWindow<R>, saved: bool) {
    let label = window.label().to_string();
    if !saved {
        cancel(&label);
        return;
    }
    if let Some(action) = resolve_one(&label) {
        complete(&app, action);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The statics are process-wide, so these run under one lock and reset
    /// first rather than racing each other.
    static GUARD: std::sync::Mutex<()> = std::sync::Mutex::new(());

    fn fresh() -> std::sync::MutexGuard<'static, ()> {
        let g = GUARD.lock().unwrap_or_else(|e| e.into_inner());
        *STATE.lock() = State::default();
        g
    }

    fn labels(names: &[&str]) -> Vec<String> {
        names.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn only_the_first_close_owns_a_window() {
        let _g = fresh();
        let target = labels(&["a"]);
        assert!(claim(&target, Action::CloseWindow("a".into())).is_some());
        // Hammering the close button must not stack a second timeout.
        assert!(claim(&target, Action::CloseWindow("a".into())).is_none());
        assert!(in_progress("a"));
    }

    #[test]
    fn a_cancelled_close_can_be_retried() {
        let _g = fresh();
        let target = labels(&["a"]);
        claim(&target, Action::CloseWindow("a".into())).unwrap();
        cancel("a");
        assert!(!in_progress("a"));
        assert!(claim(&target, Action::CloseWindow("a".into())).is_some());
    }

    #[test]
    fn cancelling_disarms_the_timeout_already_in_flight() {
        let _g = fresh();
        let target = labels(&["a"]);
        let attempt = claim(&target, Action::CloseWindow("a".into())).unwrap();
        cancel("a");
        // The sleeping thread's timeout must not act on a claim that no
        // longer exists.
        assert!(resolve_timeout(attempt).is_none());
    }

    #[test]
    fn a_later_attempt_does_not_revive_an_old_timeout() {
        let _g = fresh();
        let target = labels(&["a"]);
        let first = claim(&target, Action::CloseWindow("a".into())).unwrap();
        cancel("a");
        let second = claim(&target, Action::CloseWindow("a".into())).unwrap();

        assert_ne!(first, second);
        assert!(resolve_timeout(first).is_none());
        assert!(in_progress("a"));
        let _ = resolve_timeout(second); // consumes the real attempt, cleanly
    }

    #[test]
    fn quitting_waits_for_every_window_to_confirm() {
        let _g = fresh();
        let target = labels(&["a", "b"]);
        claim(&target, Action::ExitProcess).unwrap();

        assert!(resolve_one("a").is_none(), "b has not confirmed yet");
        assert!(matches!(resolve_one("b"), Some(Action::ExitProcess)));
    }

    #[test]
    fn one_window_failing_to_save_cancels_the_whole_quit() {
        let _g = fresh();
        let target = labels(&["a", "b", "c"]);
        claim(&target, Action::ExitProcess).unwrap();

        // a already confirmed; b then fails to save.
        assert!(resolve_one("a").is_none());
        cancel("b");

        // Every window in the attempt is released, including the one that
        // already confirmed - none of them should be stuck unable to close.
        assert!(!in_progress("a"));
        assert!(!in_progress("b"));
        assert!(!in_progress("c"));
    }

    #[test]
    fn a_window_mid_close_blocks_a_concurrent_quit_over_it() {
        let _g = fresh();
        claim(&labels(&["a"]), Action::CloseWindow("a".into())).unwrap();
        // Cmd+Q arriving while "a" is already closing must not steal it into
        // a second, competing attempt.
        assert!(claim(&labels(&["a", "b"]), Action::ExitProcess).is_none());
    }
}
