//! Flush before the process ends.
//!
//! Autosave debounces, so at any instant the newest edit may still be in
//! memory. Closing the window or quitting used to tear the process down
//! immediately, taking that edit with it - the one outcome the product exists
//! to prevent. Every exit now runs through here instead: the frontend is asked
//! to write, and only then does the app exit.
//!
//! Two ways out of the wait, both deliberate. A frontend that reports the write
//! failed cancels the exit, so the user keeps the window (and the round) and can
//! do something about the full disk or ejected drive. A frontend that never
//! answers at all is overridden after a timeout, because a hung renderer must
//! not trap someone in an app they asked to close.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Duration;

use tauri::{AppHandle, Emitter, Runtime};

/// How long the window stays open waiting for the frontend to confirm.
const FLUSH_TIMEOUT: Duration = Duration::from_secs(3);

/// True between asking the frontend to flush and actually exiting.
static QUITTING: AtomicBool = AtomicBool::new(false);

/// Invalidates a pending timeout when an exit is cancelled, so a later attempt
/// is not killed by the previous attempt's deadline.
static ATTEMPT: AtomicU64 = AtomicU64::new(0);

/// Claim the exit. `Some(attempt)` means this call owns it and should ask the
/// frontend to flush; `None` means one is already in flight, so hammering Cmd+Q
/// cannot stack timeouts.
fn claim() -> Option<u64> {
    if QUITTING.swap(true, Ordering::SeqCst) {
        return None;
    }
    Some(ATTEMPT.load(Ordering::SeqCst))
}

/// Whether the deadline belonging to `attempt` should still fire.
fn deadline_lives(attempt: u64) -> bool {
    ATTEMPT.load(Ordering::SeqCst) == attempt && QUITTING.load(Ordering::SeqCst)
}

/// Abandon the exit, invalidating any deadline already in flight.
fn cancel() {
    ATTEMPT.fetch_add(1, Ordering::SeqCst);
    QUITTING.store(false, Ordering::SeqCst);
}

/// True once an exit is under way, so the close handler stops intercepting.
pub fn in_progress() -> bool {
    QUITTING.load(Ordering::SeqCst)
}

/// Begin an orderly exit.
pub fn request<R: Runtime>(app: &AppHandle<R>) {
    let Some(attempt) = claim() else { return };
    let _ = app.emit("app:flush", ());

    let handle = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(FLUSH_TIMEOUT);
        if deadline_lives(attempt) {
            handle.exit(0);
        }
    });
}

/// The frontend's answer. `saved` false keeps the app open with the round
/// intact; the frontend is responsible for saying why.
#[tauri::command]
pub fn finish_quit<R: Runtime>(app: AppHandle<R>, saved: bool) {
    if saved {
        app.exit(0);
    } else {
        cancel();
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
        QUITTING.store(false, Ordering::SeqCst);
        g
    }

    #[test]
    fn only_the_first_request_owns_the_exit() {
        let _g = fresh();
        assert!(claim().is_some());
        // Hammering Cmd+Q must not stack a second timeout thread.
        assert!(claim().is_none());
        assert!(in_progress());
    }

    #[test]
    fn a_cancelled_exit_can_be_retried() {
        let _g = fresh();
        assert!(claim().is_some());
        cancel();
        assert!(!in_progress());
        assert!(claim().is_some(), "a second attempt must be able to start");
    }

    #[test]
    fn cancelling_disarms_the_deadline_already_in_flight() {
        let _g = fresh();
        let attempt = claim().expect("first claim owns the exit");
        assert!(deadline_lives(attempt));

        // The frontend reported it could not save.
        cancel();

        // The sleeping thread must not exit the app out from under the user
        // who is still fixing their full disk.
        assert!(!deadline_lives(attempt));
    }

    #[test]
    fn a_later_attempt_does_not_revive_an_old_deadline() {
        let _g = fresh();
        let first = claim().expect("first claim owns the exit");
        cancel();
        let second = claim().expect("second attempt starts");

        assert_ne!(first, second);
        assert!(!deadline_lives(first));
        assert!(deadline_lives(second));
    }
}
