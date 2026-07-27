//! The iroh endpoint behind the PeerLink port.
//!
//! Everything about what the peers say to each other lives in TypeScript and
//! is proven against an in-memory transport. This module only carries bytes:
//! it binds an endpoint, accepts and dials connections on one ALPN, and moves
//! newline-delimited JSON in both directions.
//!
//! The endpoint is built from `presets::Minimal`, never `presets::N0`. N0 adds
//! a Pkarr publisher and a DNS lookup pointed at n0.computer, which would
//! publish this install to a public registry on every launch. Minimal sets the
//! crypto provider and nothing else, so an idle ebb says nothing about itself
//! anywhere. mDNS is added only when asked for, and the DNS and DHT lookup
//! crates are not dependencies at all.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::Arc;

use iroh::endpoint::{presets, Connection, TransportAddrUsage};
use iroh::{Endpoint, EndpointAddr, EndpointId, RelayMode, SecretKey};
use iroh_mdns_address_lookup::MdnsAddressLookup;
use parking_lot::Mutex;
use serde::Serialize;
use tauri::{AppHandle, Emitter, State};
use tokio::io::{AsyncBufReadExt, BufReader};
use tokio::sync::mpsc;

/// One round's worth of protocol. Bumped only for a change an older build
/// cannot read, which the handshake refuses by version rather than by parse.
pub const ALPN: &[u8] = b"ebb/flow/1";

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct PeerEvent {
    conn_id: String,
    endpoint_id: String,
    connection_type: String,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct MessageEvent {
    conn_id: String,
    payload: String,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct ClosedEvent {
    conn_id: String,
}

/// What a dial hands back. The connection type rides along rather than
/// arriving as an event, so the caller never has to correlate the two.
#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DialResult {
    conn_id: String,
    connection_type: String,
}

struct Live {
    runtime: tokio::runtime::Runtime,
    endpoint: Endpoint,
    endpoint_id: String,
    conns: Arc<Mutex<HashMap<String, mpsc::UnboundedSender<String>>>>,
    next_conn: Arc<Mutex<u64>>,
}

#[derive(Default)]
pub struct CollabState {
    live: Mutex<Option<Live>>,
}

fn relay_mode(relay: bool) -> RelayMode {
    // Off restricts a session to links the two machines can make themselves.
    if relay {
        RelayMode::Default
    } else {
        RelayMode::Disabled
    }
}

/// The identity file, kept beside the config file.
fn identity_path() -> Option<PathBuf> {
    crate::config::config_dir().map(|dir| dir.join("identity.key"))
}

/// This install's long-lived secret key, minted on first use.
///
/// Saved contacts and a host's known-peer list are keyed by EndpointId, which
/// is this key's public half, so a key drawn fresh on every bind would turn a
/// saved partner back into a stranger after a restart.
///
/// None means the key could not be stored, which a read-only config directory
/// is enough to cause. That is not fatal: the endpoint mints its own and the
/// session runs on an identity that lasts one process.
fn load_or_create_secret_key() -> Option<SecretKey> {
    secret_key_at(&identity_path()?)
}

fn secret_key_at(path: &Path) -> Option<SecretKey> {
    let stored = std::fs::read_to_string(path)
        .ok()
        .and_then(|text| decode_hex(text.trim()));
    if let Some(bytes) = stored {
        return Some(SecretKey::from_bytes(&bytes));
    }
    // Unreadable, truncated and malformed are one case: whatever is on disk is
    // not a key, so a new one takes its place.
    let key = SecretKey::generate();
    write_identity(path, &encode_hex(&key.to_bytes())).ok()?;
    Some(key)
}

/// Writes the key where only this account can read it. Anyone holding these
/// bytes can present themselves as this install to a saved contact.
fn write_identity(path: &Path, hex: &str) -> std::io::Result<()> {
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    #[cfg(unix)]
    {
        use std::io::Write;
        use std::os::unix::fs::{OpenOptionsExt, PermissionsExt};
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .mode(0o600)
            .open(path)?;
        // `mode` applies only when the open creates the file, so a file left by
        // an earlier run would otherwise keep whatever permissions it has.
        file.set_permissions(std::fs::Permissions::from_mode(0o600))?;
        file.write_all(hex.as_bytes())
    }
    #[cfg(not(unix))]
    {
        std::fs::write(path, hex)
    }
}

fn encode_hex(bytes: &[u8; 32]) -> String {
    use std::fmt::Write;
    bytes
        .iter()
        .fold(String::with_capacity(64), |mut out, byte| {
            let _ = write!(out, "{byte:02x}");
            out
        })
}

fn decode_hex(text: &str) -> Option<[u8; 32]> {
    if text.len() != 64 {
        return None;
    }
    let mut bytes = [0u8; 32];
    for (slot, pair) in bytes.iter_mut().zip(text.as_bytes().chunks_exact(2)) {
        let high = char::from(pair[0]).to_digit(16)?;
        let low = char::from(pair[1]).to_digit(16)?;
        *slot = (high * 16 + low) as u8;
    }
    Some(bytes)
}

/// Pumps one connection's bidirectional stream in both directions.
///
/// The dialling side must write before the listening side's `accept_bi` can
/// return, which is exactly the protocol's shape: the guest speaks first with
/// its hello.
fn spawn_conn(
    app: AppHandle,
    state_conns: Arc<Mutex<HashMap<String, mpsc::UnboundedSender<String>>>>,
    conn_id: String,
    conn: Connection,
    send: iroh::endpoint::SendStream,
    recv: iroh::endpoint::RecvStream,
) {
    let (tx, mut rx) = mpsc::unbounded_channel::<String>();
    state_conns.lock().insert(conn_id.clone(), tx);

    // Outbound: one JSON object per line.
    let mut send = send;
    let writer_id = conn_id.clone();
    tokio::spawn(async move {
        while let Some(line) = rx.recv().await {
            if send.write_all(line.as_bytes()).await.is_err() {
                break;
            }
            if send.write_all(b"\n").await.is_err() {
                break;
            }
        }
        let _ = send.finish();
        drop(writer_id);
    });

    // Inbound: each line is one wire message for the webview.
    let reader_id = conn_id.clone();
    tokio::spawn(async move {
        let mut lines = BufReader::new(recv).lines();
        while let Ok(Some(line)) = lines.next_line().await {
            if line.is_empty() {
                continue;
            }
            let _ = app.emit(
                "collab:message",
                MessageEvent {
                    conn_id: reader_id.clone(),
                    payload: line,
                },
            );
        }
        conn.closed().await;
        state_conns.lock().remove(&reader_id);
        let _ = app.emit("collab:closed", ClosedEvent { conn_id: reader_id });
    });
}

/// Direct or relayed, as the endpoint actually observes it.
///
/// An address is only direct when it is an IP path that is actually in use.
/// Anything else reports relayed: a relay disclosure that has to guess must
/// guess in the direction that over-discloses, never under.
async fn connection_type(endpoint: &Endpoint, remote: EndpointId) -> String {
    match endpoint.remote_info(remote).await {
        Some(info)
            if info
                .addrs()
                .any(|a| matches!(a.usage(), TransportAddrUsage::Active) && a.addr().is_ip()) =>
        {
            "direct".to_string()
        }
        _ => "relayed".to_string(),
    }
}

#[tauri::command]
pub fn collab_start(
    app: AppHandle,
    state: State<'_, CollabState>,
    relay: bool,
    mdns: bool,
) -> Result<String, String> {
    let mut held = state.live.lock();
    if let Some(live) = held.as_ref() {
        return Ok(live.endpoint_id.clone());
    }

    let runtime = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|e| format!("Could not start the collaboration runtime: {e}"))?;

    let endpoint = runtime.block_on(async {
        // Minimal, never N0: see the module comment.
        let mut builder = Endpoint::builder(presets::Minimal)
            .alpns(vec![ALPN.to_vec()])
            .relay_mode(relay_mode(relay));
        if let Some(key) = load_or_create_secret_key() {
            builder = builder.secret_key(key);
        }
        let endpoint = builder
            .bind()
            .await
            .map_err(|e| format!("Could not bind an endpoint: {e}"))?;

        if mdns {
            let lookup = MdnsAddressLookup::builder()
                .build(endpoint.id())
                .map_err(|e| format!("Could not start local discovery: {e}"))?;
            if let Ok(services) = endpoint.address_lookup() {
                services.add(lookup);
            }
        }
        Ok::<Endpoint, String>(endpoint)
    })?;

    let endpoint_id = endpoint.id().to_string();
    let conns: Arc<Mutex<HashMap<String, mpsc::UnboundedSender<String>>>> =
        Arc::new(Mutex::new(HashMap::new()));
    let next_conn = Arc::new(Mutex::new(0u64));

    // Accept loop.
    {
        let app = app.clone();
        let endpoint = endpoint.clone();
        let conns = conns.clone();
        let next_conn = next_conn.clone();
        runtime.spawn(async move {
            while let Some(incoming) = endpoint.accept().await {
                let Ok(conn) = incoming.await else { continue };
                let Ok((send, recv)) = conn.accept_bi().await else {
                    continue;
                };
                let remote = conn.remote_id();
                let conn_id = {
                    let mut n = next_conn.lock();
                    *n += 1;
                    format!("c{n}")
                };
                let _ = app.emit(
                    "collab:peer",
                    PeerEvent {
                        conn_id: conn_id.clone(),
                        endpoint_id: remote.to_string(),
                        connection_type: connection_type(&endpoint, remote).await,
                    },
                );
                spawn_conn(app.clone(), conns.clone(), conn_id, conn, send, recv);
            }
        });
    }

    *held = Some(Live {
        runtime,
        endpoint,
        endpoint_id: endpoint_id.clone(),
        conns,
        next_conn,
    });
    Ok(endpoint_id)
}

#[tauri::command]
pub fn collab_dial(
    app: AppHandle,
    state: State<'_, CollabState>,
    endpoint_id: String,
) -> Result<DialResult, String> {
    let held = state.live.lock();
    let live = held.as_ref().ok_or("No collaboration session is running")?;

    let remote = EndpointId::from_str(&endpoint_id).map_err(|_| "Not an endpoint id".to_string())?;
    let endpoint = live.endpoint.clone();
    let conn = live
        .runtime
        .block_on(async move { endpoint.connect(EndpointAddr::new(remote), ALPN).await })
        .map_err(|e| format!("Could not reach that peer: {e}"))?;

    let (send, recv) = live
        .runtime
        .block_on(async { conn.open_bi().await })
        .map_err(|e| format!("Could not open a stream: {e}"))?;

    let conn_id = {
        let mut n = live.next_conn.lock();
        *n += 1;
        format!("c{n}")
    };
    let kind = live
        .runtime
        .block_on(async { connection_type(&live.endpoint, remote).await });
    // No collab:peer event here. The dialler already holds this connection,
    // and announcing it would race the return of this very call: the webview
    // could see the event before it learns the id and mistake its own dial
    // for an inbound peer.
    let conns = live.conns.clone();
    let handle = app.clone();
    let id = conn_id.clone();
    live.runtime
        .spawn(async move { spawn_conn(handle, conns, id, conn, send, recv) });
    Ok(DialResult {
        conn_id,
        connection_type: kind,
    })
}

#[tauri::command]
pub fn collab_send(
    state: State<'_, CollabState>,
    conn_id: String,
    payload: String,
) -> Result<(), String> {
    let held = state.live.lock();
    let live = held.as_ref().ok_or("No collaboration session is running")?;
    let conns = live.conns.lock();
    let tx = conns.get(&conn_id).ok_or("That peer is gone")?;
    tx.send(payload).map_err(|_| "That peer is gone".to_string())
}

#[tauri::command]
pub fn collab_close(state: State<'_, CollabState>, conn_id: String) -> Result<(), String> {
    let held = state.live.lock();
    let live = held.as_ref().ok_or("No collaboration session is running")?;
    live.conns.lock().remove(&conn_id);
    Ok(())
}

#[tauri::command]
pub fn collab_stop(state: State<'_, CollabState>) -> Result<(), String> {
    let Some(live) = state.live.lock().take() else {
        return Ok(());
    };
    live.conns.lock().clear();
    let endpoint = live.endpoint.clone();
    live.runtime.block_on(async move { endpoint.close().await });
    // Dropping the runtime here would block this thread on its workers, so it
    // is handed to one that is allowed to wait.
    std::thread::spawn(move || drop(live.runtime));
    Ok(())
}

/// Strips what a hostname carries for the network and a debater does not: the
/// mDNS suffix, and any domain past the first label. `smodi-mbp.local` and
/// `smodi-mbp.tourney.lan` are both `smodi-mbp`.
fn short_host(raw: &str) -> String {
    raw.trim()
        .trim_end_matches('.')
        .split('.')
        .next()
        .unwrap_or("")
        .to_string()
}

/// What this machine calls itself, for the display name a shared round
/// carries. Read from the `hostname` binary every desktop platform ships,
/// rather than pulling in a crate for one string. Empty when it cannot be
/// read, which the caller shows as no default rather than as an error.
#[tauri::command]
pub fn machine_name() -> String {
    let Ok(out) = std::process::Command::new("hostname").output() else {
        return String::new();
    };
    if !out.status.success() {
        return String::new();
    }
    short_host(&String::from_utf8_lossy(&out.stdout))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_alpn_names_the_protocol_version() {
        assert_eq!(ALPN, b"ebb/flow/1");
    }

    #[test]
    fn relay_off_disables_relays_outright() {
        assert_eq!(relay_mode(false), RelayMode::Disabled);
        assert_eq!(relay_mode(true), RelayMode::Default);
    }

    #[test]
    fn a_host_keeps_its_first_label_and_nothing_else() {
        assert_eq!(short_host("smodi-mbp.local\n"), "smodi-mbp");
        assert_eq!(short_host("smodi-mbp.tourney.lan"), "smodi-mbp");
        assert_eq!(short_host("  smodi-mbp  "), "smodi-mbp");
        assert_eq!(short_host(""), "");
    }
}

/// The stored key behind a stable EndpointId.
#[cfg(test)]
mod identity {
    use super::*;

    const KNOWN: &str = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";

    /// A path under a directory of this test's own, never the real config dir.
    fn scratch(tag: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("ebb-collab-{tag}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        let _ = std::fs::remove_file(&dir);
        dir.join("identity.key")
    }

    /// The wiring `collab_start` uses: the stored key, or the endpoint's own
    /// when there is none.
    async fn bound_with(path: &Path) -> Endpoint {
        let mut builder = Endpoint::builder(presets::Minimal)
            .alpns(vec![ALPN.to_vec()])
            .relay_mode(RelayMode::Disabled);
        if let Some(key) = secret_key_at(path) {
            builder = builder.secret_key(key);
        }
        builder.bind().await.expect("bind")
    }

    /// A contact is saved under an EndpointId, which is the public half of the
    /// stored key, so two binds from one identity file are one peer.
    #[tokio::test]
    async fn rebinding_keeps_the_endpoint_id() {
        let path = scratch("rebind");

        let first = bound_with(&path).await;
        let id = first.id();
        drop(first);
        let second = bound_with(&path).await;

        assert_eq!(second.id(), id);
    }

    #[test]
    fn a_stored_key_is_read_back_byte_for_byte() {
        let path = scratch("stored");
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, format!("{KNOWN}\n")).unwrap();

        let key = secret_key_at(&path).expect("the stored key");
        assert_eq!(encode_hex(&key.to_bytes()), KNOWN);
    }

    #[test]
    fn a_minted_key_is_the_same_key_on_the_next_run() {
        let path = scratch("minted");

        let first = secret_key_at(&path).expect("a minted key");
        let second = secret_key_at(&path).expect("the same key again");

        assert_eq!(first.to_bytes(), second.to_bytes());
        assert_eq!(std::fs::read_to_string(&path).unwrap().len(), 64);
    }

    #[test]
    fn a_malformed_file_yields_a_fresh_key_rather_than_an_error() {
        for (tag, junk) in [
            ("empty", ""),
            ("short", "abcd"),
            ("prose", "not a key"),
            ("nonhex", "zz23456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"),
        ] {
            let path = scratch(tag);
            std::fs::create_dir_all(path.parent().unwrap()).unwrap();
            std::fs::write(&path, junk).unwrap();

            let key = secret_key_at(&path).expect("a fresh key");
            assert_eq!(
                std::fs::read_to_string(&path).unwrap(),
                encode_hex(&key.to_bytes()),
                "{tag}"
            );
        }
    }

    #[cfg(unix)]
    #[test]
    fn the_stored_key_is_readable_only_by_its_owner() {
        use std::os::unix::fs::PermissionsExt;
        let path = scratch("mode");
        secret_key_at(&path).expect("a minted key");

        let mode = std::fs::metadata(&path).unwrap().permissions().mode();
        assert_eq!(mode & 0o777, 0o600);
    }

    #[test]
    fn a_config_directory_that_cannot_be_written_is_not_fatal() {
        let blocked = scratch("blocked");
        std::fs::write(blocked.parent().unwrap(), "a file where a directory would go").unwrap();

        assert!(secret_key_at(&blocked).is_none());
    }

    #[test]
    fn hex_decoding_takes_exactly_thirty_two_bytes() {
        let bytes = decode_hex(KNOWN).expect("64 hex characters");
        assert_eq!(&bytes[..4], &[0x01, 0x23, 0x45, 0x67]);
        assert_eq!(encode_hex(&bytes), KNOWN);
        assert!(decode_hex(&KNOWN[..62]).is_none());
        assert!(decode_hex(&format!("{KNOWN}00")).is_none());
    }
}

#[cfg(test)]
mod preset_guard {
    /// The difference between Minimal and N0 is one word, and choosing wrong
    /// publishes this install to a public DNS registry on every launch. That
    /// is not something a reviewer should have to catch by eye.
    #[test]
    fn the_endpoint_is_never_built_from_the_n0_preset() {
        // Assembled at runtime so this test's own source does not contain the
        // string it is scanning for.
        let banned = concat!("presets", "::N0");
        let source = include_str!("collab.rs");
        for line in source.lines() {
            let code = line.trim_start();
            if code.starts_with("//") || code.contains("concat!") {
                continue;
            }
            assert!(!code.contains(banned), "that preset publishes to DNS: {line}");
        }
        assert!(source.contains("presets::Minimal"));
    }
}

/// Real endpoints over real QUIC on loopback.
///
/// Everything above the socket is proven against an in-memory transport, so
/// what is left to prove is that this configuration actually carries bytes:
/// the Minimal preset with relays disabled and no discovery at all, which is
/// exactly the shape an isolated room gets. Addresses are passed directly, so
/// nothing is published anywhere for these to find each other.
#[cfg(test)]
mod loopback {
    use super::*;
    use tokio::io::AsyncBufReadExt;

    async fn bound() -> Endpoint {
        Endpoint::builder(presets::Minimal)
            .alpns(vec![ALPN.to_vec()])
            .relay_mode(RelayMode::Disabled)
            .bind()
            .await
            .expect("bind")
    }

    #[tokio::test]
    async fn two_endpoints_exchange_a_wire_message() {
        let host = bound().await;
        let guest = bound().await;
        let host_addr = host.addr();

        // Both halves run in one scope. Dropping either the endpoint or the
        // connection closes it, so a spawned task that returns early tears the
        // link down before the other side has read its reply.
        let listener = async {
            let incoming = host.accept().await.expect("incoming");
            let conn = incoming.await.expect("accept");
            let (mut send, recv) = conn.accept_bi().await.expect("accept_bi");
            let mut lines = BufReader::new(recv).lines();
            let hello = lines.next_line().await.expect("read").expect("a line");
            send.write_all(b"{\"type\":\"helloAck\",\"ok\":true}\n")
                .await
                .expect("write");
            send.finish().expect("finish");
            // Held open until the dialler hangs up, which it does once it has
            // the reply.
            conn.closed().await;
            hello
        };

        let dialler = async {
            let conn = guest.connect(host_addr, ALPN).await.expect("connect");
            let (mut send, recv) = conn.open_bi().await.expect("open_bi");
            send.write_all(b"{\"type\":\"hello\",\"protocol\":1}\n")
                .await
                .expect("write");
            let mut lines = BufReader::new(recv).lines();
            let ack = lines.next_line().await.expect("read").expect("a line");
            conn.close(0u32.into(), b"done");
            ack
        };

        let (hello, ack) = tokio::join!(listener, dialler);
        assert_eq!(hello, "{\"type\":\"hello\",\"protocol\":1}");
        assert_eq!(ack, "{\"type\":\"helloAck\",\"ok\":true}");
    }

    #[tokio::test]
    async fn a_connection_on_another_protocol_is_refused() {
        let host = bound().await;
        let guest = bound().await;
        let addr = host.addr();
        let listener = host.clone();
        tokio::spawn(async move {
            if let Some(incoming) = listener.accept().await {
                let _ = incoming.await;
            }
        });
        assert!(guest.connect(addr, b"someone/else/1").await.is_err());
        drop(host);
    }
}
