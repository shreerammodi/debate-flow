//! Ebb desktop shell.
//!
//! The shell stays deliberately thin: all app logic lives in the React
//! frontend (the same `src/` that powers the web build). Rust owns only window
//! creation, the native menu, lifecycle guards, and (later) the updater.

mod bridge;
mod config;
mod flowfile;
mod menu;

use tauri::{Emitter, Manager};

/// `[os, arch]` of the running binary, e.g. `["macos", "aarch64"]`. The webview
/// user agent can't be trusted for either (macOS reports "Intel" on Apple
/// Silicon), so the values come from the compiled target.
#[tauri::command]
fn system_info() -> [&'static str; 2] {
    [std::env::consts::OS, std::env::consts::ARCH]
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let builder = tauri::Builder::default();

    // Must precede every other plugin. A second launch - a double-clicked .ebb
    // on Windows or Linux - hands its argv to the running process and focuses
    // the existing window, instead of starting a rival copy that would autosave
    // over the same files.
    #[cfg(desktop)]
    let builder = builder.plugin(tauri_plugin_single_instance::init(|app, argv, _cwd| {
        if let Some(window) = app.get_webview_window("main") {
            let _ = window.set_focus();
        }
        for path in flowfile::flow_paths_in(&argv) {
            flowfile::request_open(app, path);
        }
    }));

    builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(flowfile::PendingOpen::default())
        .setup(|app| {
            // Signed updater + relaunch (desktop only). Policy (when to
            // download, install only on user confirmation) lives in the JS
            // layer; these plugins just expose the verified
            // check/download/install/relaunch primitives it drives.
            #[cfg(desktop)]
            {
                app.handle()
                    .plugin(tauri_plugin_updater::Builder::new().build())?;
                app.handle().plugin(tauri_plugin_process::init())?;
            }

            // Mirror settings to a plain-text config file and watch it for
            // external edits (desktop only; see `config.rs`).
            #[cfg(desktop)]
            config::init(app.handle());

            // Loopback endpoint CardMirror sends extracted cards to, and the
            // broker for our own outbound calls (desktop only; see
            // `bridge.rs`). A failed bind costs the integration, not the app.
            #[cfg(desktop)]
            bridge::start(app.handle());

            // Install the native menu; accelerators follow the JS keymap via
            // rebuild_menu (see menu.rs).
            let handle = app.handle();
            let app_menu = menu::build(handle, &std::collections::HashMap::new())?;
            app.set_menu(app_menu)?;

            // A .ebb opened from the file manager at launch. macOS delivers it
            // as RunEvent::Opened instead; both routes buffer until the
            // frontend drains them.
            for path in flowfile::flow_paths_in(&std::env::args().collect::<Vec<_>>()) {
                flowfile::request_open(handle, path);
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            bridge::bridge_reply,
            bridge::cardmirror_insert,
            bridge::cardmirror_jump,
            bridge::cardmirror_status,
            config::read_config,
            config::write_config,
            flowfile::create_flow_file,
            flowfile::drain_pending_open,
            flowfile::flow_paths,
            flowfile::read_flow_file,
            flowfile::read_recents,
            flowfile::write_flow_file,
            flowfile::write_recents,
            menu::rebuild_menu,
            system_info
        ])
        // Quit is the single deliberate exit; it routes here and exits directly,
        // bypassing the close guard below. Every other menu item carries a JS
        // CommandId, which we hand to the frontend to run (see useDesktopMenu).
        .on_menu_event(|app, event| {
            let id = event.id().0.as_str();
            if id == menu::QUIT_ID {
                app.exit(0);
            } else {
                let _ = app.emit("menu:command", id.to_string());
            }
        })
        .build(tauri::generate_context!())
        .expect("error while running Ebb")
        // The session handshake advertises a port that dies with the process,
        // so it has to go on the way out; the identity file stays.
        .run(|app, event| {
            // Only the macOS arm below reads the handle.
            #[cfg(not(target_os = "macos"))]
            let _ = app;
            match event {
                tauri::RunEvent::Exit => bridge::remove_session(),
                // macOS "Open With" and double-click, at launch and while
                // already running.
                #[cfg(target_os = "macos")]
                tauri::RunEvent::Opened { urls } => {
                    for url in urls {
                        if let Ok(path) = url.to_file_path() {
                            flowfile::request_open(app, path.to_string_lossy().into_owned());
                        }
                    }
                }
                _ => {}
            }
        });
}
