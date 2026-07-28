//! Runtime window creation and focus tracking.
//!
//! ebb has no privileged "main" window: every dashboard and every flow editor
//! is a fully independent window, built here rather than declared in
//! tauri.conf.json (whose `app.windows` list is empty) so a cold launch, a
//! second launch, and Mod+N can each decide how many windows to open and
//! where each one starts. Opening a flow always creates a new window rather
//! than steering an existing one, so a debater who pulls up a second flow
//! keeps the first exactly as it was.
//!
//! The one piece of state shared across windows is which one is currently
//! focused: a menu accelerator and the CardMirror bridge both need "the
//! window the user is looking at", and neither Tauri callback hands that to
//! us directly.

use std::sync::atomic::{AtomicU64, Ordering};

use parking_lot::Mutex;
use tauri::{AppHandle, Manager, Runtime, WebviewUrl, WebviewWindow, WebviewWindowBuilder};

/// Chrome shared by every window; mirrors the single static entry
/// tauri.conf.json declared before windows became dynamic.
const TITLE: &str = "ebb";
const WIDTH: f64 = 1280.0;
const HEIGHT: f64 = 800.0;
const MIN_HEIGHT: f64 = 600.0;

static NEXT_ID: AtomicU64 = AtomicU64::new(0);

/// The most recently focused window's label.
static FOCUSED: Mutex<Option<String>> = Mutex::new(None);

pub fn note_focus<R: Runtime>(window: &tauri::Window<R>) {
    *FOCUSED.lock() = Some(window.label().to_string());
}

/// Clears the focus record if it still points at `label`, so a destroyed
/// window is never handed back as a stale target.
pub fn note_close(label: &str) {
    let mut focused = FOCUSED.lock();
    if focused.as_deref() == Some(label) {
        *focused = None;
    }
}

/// The window to route a single-recipient action (a menu accelerator, a
/// CardMirror request) at: the last-focused one, or - if focus was never
/// observed, e.g. the very first event after launch - any other open window.
pub fn target_window<R: Runtime>(app: &AppHandle<R>) -> Option<WebviewWindow<R>> {
    let label = FOCUSED.lock().clone();
    label
        .and_then(|l| app.get_webview_window(&l))
        .or_else(|| app.webview_windows().values().next().cloned())
}

fn build<R: Runtime>(app: &AppHandle<R>, url: WebviewUrl) -> tauri::Result<WebviewWindow<R>> {
    let label = format!("win-{}", NEXT_ID.fetch_add(1, Ordering::SeqCst));
    WebviewWindowBuilder::new(app, label, url)
        .title(TITLE)
        .inner_size(WIDTH, HEIGHT)
        .min_inner_size(0.0, MIN_HEIGHT)
        .resizable(true)
        .fullscreen(false)
        .disable_drag_drop_handler()
        .build()
}

/// Opens a new window on the dashboard.
pub fn open_dashboard<R: Runtime>(app: &AppHandle<R>) -> tauri::Result<WebviewWindow<R>> {
    build(app, WebviewUrl::App("index.html".into()))
}

/// Opens a new window on the given flow.
pub fn open_flow<R: Runtime>(app: &AppHandle<R>, path: &str) -> tauri::Result<WebviewWindow<R>> {
    // Reuses the `Url` query-pair encoder to match the percent-encoding
    // `encodeURIComponent` produces on the frontend (see flowNav.ts's
    // flowRouteFor); the scheme and host here are thrown away immediately.
    let mut qs = tauri::Url::parse("app://ebb").expect("static URL parses");
    qs.query_pairs_mut().append_pair("path", path);
    let target = format!("flow/?{}", qs.query().unwrap_or_default());
    build(app, WebviewUrl::App(target.into()))
}

/// Opens a new dashboard window. The JS side of `window.new` (Mod+N, the
/// Window menu, the command palette).
#[tauri::command]
pub fn new_window<R: Runtime>(app: AppHandle<R>) -> Result<(), String> {
    open_dashboard(&app).map(|_| ()).map_err(|e| e.to_string())
}

// --- Cold-launch bootstrap ------------------------------------------------------
//
// A .ebb opened from the file manager reaches Rust as argv (Windows/Linux,
// available synchronously before any window exists, so setup() can just
// build the right window directly) or as RunEvent::Opened (macOS, which can
// only be observed once the run loop is already pumping - after setup() has
// already had to decide whether to open a dashboard). Rather than race that
// delivery, setup() opens its default dashboard as usual and marks it here;
// if a path then turns out to have been requested by the very same launch,
// it adopts that still-blank window instead of leaving a redundant one
// beside a new flow window. Once settled - adopted, or drained empty - the
// window is ordinary again and never adopts a later, unrelated open.

struct Bootstrap {
    /// The dashboard window setup() created with nothing requested. A
    /// permanent identity, so a drain from a different (e.g. Mod+N-opened)
    /// dashboard can tell it is not the one being asked for.
    window: Option<String>,
    /// True once a path has been buffered for adoption, or the window has
    /// drained with nothing pending - either way, adoption is no longer
    /// available, so neither a second simultaneously-opened file nor a
    /// stray late arrival can overwrite or reclaim it.
    settled: bool,
    pending: Option<String>,
}

static BOOTSTRAP: Mutex<Bootstrap> = Mutex::new(Bootstrap {
    window: None,
    settled: false,
    pending: None,
});

/// Marks `label` as the window that can still adopt a same-launch file open.
/// Called at most once, right after setup() opens a dashboard with nothing
/// requested.
pub fn mark_bootstrap(label: &str) {
    BOOTSTRAP.lock().window = Some(label.to_string());
}

/// Opens `path`, adopting the bootstrap window in its place if one is still
/// available, otherwise opening a brand new window.
pub fn adopt_or_open<R: Runtime>(app: &AppHandle<R>, path: &str) {
    let mut b = BOOTSTRAP.lock();
    if !b.settled && b.window.is_some() {
        b.pending = Some(path.to_string());
        b.settled = true;
        return;
    }
    drop(b);
    let _ = open_flow(app, path);
}

/// Called once by the dashboard's own frontend on mount. Returns the path
/// waiting for this exact window, if any - `None` both when nothing is
/// pending and when `window` is a different, unrelated dashboard.
#[tauri::command]
pub fn drain_boot_open<R: Runtime>(window: WebviewWindow<R>) -> Option<String> {
    let mut b = BOOTSTRAP.lock();
    if b.window.as_deref() != Some(window.label()) {
        return None;
    }
    b.settled = true;
    b.pending.take()
}
