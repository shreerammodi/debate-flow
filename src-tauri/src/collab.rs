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
use tokio::io::{AsyncBufRead, AsyncBufReadExt, AsyncReadExt, BufReader};
use tokio::sync::mpsc;

/// One round's worth of protocol. Bumped only for a change an older build
/// cannot read, which the handshake refuses by version rather than by parse.
pub const ALPN: &[u8] = b"ebb/flow/1";

/// The most one line of wire JSON may be. A connection is read from the moment
/// `accept_bi` returns, which is before the handshake has admitted anyone, so
/// this is the only thing bounding what a peer that merely opened the ALPN can
/// make this process allocate. The cap is the loopback bridge's, which is the
/// standard a peer link is held to.
const MAX_LINE: usize = 4 * 1024 * 1024;

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

/// One thing that happened on a connection, on its way to the webview.
enum Event {
    Peer(PeerEvent),
    Message(MessageEvent),
    Closed(ClosedEvent),
}

/// Where those events go. The `AppHandle` is the one that ships; the seam is
/// what lets the pump be driven with no window in front of it.
trait Events: Send + Sync + 'static {
    fn emit(&self, event: Event);
}

impl Events for AppHandle {
    fn emit(&self, event: Event) {
        let _ = match event {
            Event::Peer(e) => Emitter::emit(self, "collab:peer", e),
            Event::Message(e) => Emitter::emit(self, "collab:message", e),
            Event::Closed(e) => Emitter::emit(self, "collab:closed", e),
        };
    }
}

/// What a dial hands back. The connection type rides along rather than
/// arriving as an event, so the caller never has to correlate the two.
#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DialResult {
    conn_id: String,
    connection_type: String,
}

/// One live connection: what to write to it, and everything that has to go
/// when it closes.
struct Conn {
    tx: mpsc::UnboundedSender<String>,
    conn: Connection,
    writer: tokio::task::JoinHandle<()>,
    reader: tokio::task::JoinHandle<()>,
}

impl Conn {
    /// Ends the connection for real. Dropping the sender only stops the
    /// writer: the QUIC connection and the reader outlive it, so a peer the
    /// host refused would hold its connection and keep pumping messages into
    /// the webview until it chose to hang up.
    fn close(self) {
        self.conn.close(0u32.into(), b"closed");
        self.writer.abort();
        self.reader.abort();
    }
}

struct Live {
    runtime: tokio::runtime::Runtime,
    endpoint: Endpoint,
    endpoint_id: String,
    conns: Arc<Mutex<HashMap<String, Conn>>>,
    next_conn: Arc<Mutex<u64>>,
    accept: tokio::task::JoinHandle<()>,
    /// One per caller holding this endpoint. Two PeerLinks overlap by design -
    /// the idle invite listener and a round's session - and they share one
    /// bind, so the identity and the accept loop survive either of them
    /// stopping and only the last one out closes it. The count assumes one
    /// stop per start, which is what each holder does: it drops its handle
    /// before it asks.
    ///
    /// Sharing means an inbound connection reaches whoever is listening rather
    /// than the holder it was meant for. `collab:peer` goes to the webview,
    /// which is where a connection's purpose is decided.
    holders: usize,
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
    match std::fs::read(path) {
        Ok(bytes) => {
            let stored = std::str::from_utf8(&bytes)
                .ok()
                .and_then(|text| decode_hex(text.trim()));
            if let Some(bytes) = stored {
                return Some(SecretKey::from_bytes(&bytes));
            }
        }
        // A file that is there and cannot be read is a key this run cannot
        // reach, not a key that is gone. Minting over it would retire the
        // identity every peer saved as a contact, permanently, over something
        // as transient as a lock or a mode bit - so this run goes without one
        // and the next run that can read the file has it back.
        Err(e) if e.kind() != std::io::ErrorKind::NotFound => return None,
        Err(_) => {}
    }
    // Nothing there, or something that read fine and is not a key. Neither can
    // be recovered, so a new one takes its place.
    let key = SecretKey::generate();
    write_identity(path, &encode_hex(&key.to_bytes())).ok()?;
    Some(key)
}

/// Writes the key where only this account can read it, through a temp file in
/// the same directory that is synced and then renamed over the target.
///
/// Anyone holding these bytes can present themselves as this install to a
/// saved contact, so the mode is set as the file is created rather than once
/// it already holds the key. The rename is what makes the replacement atomic:
/// a truncating write interrupted halfway leaves a file that reads back as
/// junk, which costs the install the identity its peers have saved.
fn write_identity(path: &Path, hex: &str) -> std::io::Result<()> {
    let dir = path.parent().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, "the key has no directory")
    })?;
    std::fs::create_dir_all(dir)?;
    let name = path.file_name().and_then(|n| n.to_str()).ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, "the key has no filename")
    })?;

    let tmp = dir.join(format!(".{name}.tmp"));
    if let Err(e) = write_private(&tmp, hex.as_bytes()) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    if let Err(e) = std::fs::rename(&tmp, path) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    Ok(())
}

/// Creates a file this account alone can read, and makes its bytes durable
/// before anything renames them into place.
fn write_private(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    use std::io::Write;

    #[cfg(unix)]
    let mut file = {
        use std::os::unix::fs::{OpenOptionsExt, PermissionsExt};
        let file = std::fs::OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .mode(0o600)
            .open(path)?;
        // `mode` applies only when the open creates the file, so a temp file
        // left by a crashed run would otherwise keep whatever mode it has.
        file.set_permissions(std::fs::Permissions::from_mode(0o600))?;
        file
    };
    // ponytail: no ACL on Windows. Restricting the file there means building a
    // DACL and calling SetNamedSecurityInfoW, which lives in windows-sys - not
    // a dependency of this crate, and not one worth adding for a single file,
    // because %APPDATA% is per-user and already keeps other accounts out. What
    // is missing is protection from another process running as this same user,
    // which is the residual the loopback bridge's session token carries too.
    #[cfg(not(unix))]
    let mut file = std::fs::File::create(path)?;

    file.write_all(bytes)?;
    file.sync_all()
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

/// One read from a peer's stream.
enum Line {
    Text(String),
    /// The peer hung up. A partial line at the end is not a message anyone
    /// finished sending, so it goes with the connection.
    Eof,
    /// Past the cap, not UTF-8, or a broken stream. A refused line takes the
    /// connection with it rather than being truncated and parsed, because what
    /// would be left is not what the peer sent.
    Refused,
}

/// Reads one newline-delimited line, refusing anything longer than `cap`.
///
/// Reading one byte past the cap is what tells a line that ends exactly at it
/// from one that runs past it.
async fn read_capped_line<R: AsyncBufRead + Unpin>(reader: &mut R, cap: usize) -> Line {
    let mut buf = Vec::new();
    let Ok(read) = (&mut *reader)
        .take(cap as u64 + 1)
        .read_until(b'\n', &mut buf)
        .await
    else {
        return Line::Refused;
    };
    if read == 0 {
        return Line::Eof;
    }
    if buf.last() != Some(&b'\n') {
        return if read > cap { Line::Refused } else { Line::Eof };
    }
    buf.pop();
    if buf.last() == Some(&b'\r') {
        buf.pop();
    }
    match String::from_utf8(buf) {
        Ok(line) => Line::Text(line),
        Err(_) => Line::Refused,
    }
}

/// Pumps one connection's bidirectional stream in both directions.
///
/// The dialling side must write before the listening side's `accept_bi` can
/// return, which is exactly the protocol's shape: the guest speaks first with
/// its hello.
fn spawn_conn(
    events: Arc<dyn Events>,
    state_conns: Arc<Mutex<HashMap<String, Conn>>>,
    conn_id: String,
    conn: Connection,
    send: iroh::endpoint::SendStream,
    recv: iroh::endpoint::RecvStream,
) {
    let (tx, mut rx) = mpsc::unbounded_channel::<String>();

    // The map is held across both spawns so a task that ends immediately
    // cannot drop this connection before it is in there to drop.
    let mut conns = state_conns.lock();

    // Outbound: one JSON object per line.
    let mut send = send;
    let writer = tokio::spawn(async move {
        while let Some(line) = rx.recv().await {
            if send.write_all(line.as_bytes()).await.is_err() {
                break;
            }
            if send.write_all(b"\n").await.is_err() {
                break;
            }
        }
        let _ = send.finish();
    });

    // Inbound: each line is one wire message for the webview.
    let reader_id = conn_id.clone();
    let reader_conn = conn.clone();
    let reader_conns = state_conns.clone();
    let reader = tokio::spawn(async move {
        let mut recv = BufReader::new(recv);
        loop {
            match read_capped_line(&mut recv, MAX_LINE).await {
                Line::Text(line) => {
                    if line.is_empty() {
                        continue;
                    }
                    events.emit(Event::Message(MessageEvent {
                        conn_id: reader_id.clone(),
                        payload: line,
                    }));
                }
                Line::Eof => break,
                Line::Refused => {
                    reader_conn.close(1u32.into(), b"line refused");
                    break;
                }
            }
        }
        reader_conn.closed().await;
        reader_conns.lock().remove(&reader_id);
        events.emit(Event::Closed(ClosedEvent { conn_id: reader_id }));
    });

    conns.insert(
        conn_id,
        Conn {
            tx,
            conn,
            writer,
            reader,
        },
    );
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
    start(state.inner(), Arc::new(app), relay, mdns)
}

/// Binds the endpoint, or takes a share of the one already bound.
fn start(
    state: &CollabState,
    events: Arc<dyn Events>,
    relay: bool,
    mdns: bool,
) -> Result<String, String> {
    let mut held = state.live.lock();
    if let Some(live) = held.as_mut() {
        live.holders += 1;
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
    let conns: Arc<Mutex<HashMap<String, Conn>>> = Arc::new(Mutex::new(HashMap::new()));
    let next_conn = Arc::new(Mutex::new(0u64));

    let accept = {
        let events = events.clone();
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
                events.emit(Event::Peer(PeerEvent {
                    conn_id: conn_id.clone(),
                    endpoint_id: remote.to_string(),
                    connection_type: connection_type(&endpoint, remote).await,
                }));
                spawn_conn(events.clone(), conns.clone(), conn_id, conn, send, recv);
            }
        })
    };

    *held = Some(Live {
        runtime,
        endpoint,
        endpoint_id: endpoint_id.clone(),
        conns,
        next_conn,
        accept,
        holders: 1,
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
    let events: Arc<dyn Events> = Arc::new(app);
    let conns = live.conns.clone();
    let id = conn_id.clone();
    live.runtime
        .spawn(async move { spawn_conn(events, conns, id, conn, send, recv) });
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
    let held = conns.get(&conn_id).ok_or("That peer is gone")?;
    held.tx
        .send(payload)
        .map_err(|_| "That peer is gone".to_string())
}

#[tauri::command]
pub fn collab_close(state: State<'_, CollabState>, conn_id: String) -> Result<(), String> {
    close(state.inner(), &conn_id)
}

/// Hangs up on one peer, leaving the endpoint and every other peer up.
fn close(state: &CollabState, conn_id: &str) -> Result<(), String> {
    let held = state.live.lock();
    let live = held.as_ref().ok_or("No collaboration session is running")?;
    let gone = live.conns.lock().remove(conn_id);
    if let Some(conn) = gone {
        conn.close();
    }
    Ok(())
}

#[tauri::command]
pub fn collab_stop(state: State<'_, CollabState>) -> Result<(), String> {
    stop(state.inner())
}

/// Lets go of one holder's share of the endpoint.
fn stop(state: &CollabState) -> Result<(), String> {
    let last = {
        let mut held = state.live.lock();
        let Some(live) = held.as_mut() else {
            return Ok(());
        };
        live.holders = live.holders.saturating_sub(1);
        if live.holders > 0 {
            return Ok(());
        }
        held.take()
    };
    if let Some(live) = last {
        shutdown(live);
    }
    Ok(())
}

/// Closes every peer, then the accept loop, then the endpoint.
fn shutdown(live: Live) {
    for (_, conn) in live.conns.lock().drain() {
        conn.close();
    }
    live.accept.abort();
    let endpoint = live.endpoint.clone();
    live.runtime.block_on(async move { endpoint.close().await });
    // Dropping the runtime here would block this thread on its workers, so it
    // is handed to one that is allowed to wait.
    std::thread::spawn(move || drop(live.runtime));
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

    /// A peer that streams bytes and never a newline is not sending a message,
    /// it is asking for an allocation. The reader stops one byte past the cap
    /// and refuses what it has rather than following the stream.
    #[tokio::test]
    async fn a_line_is_refused_at_the_cap_rather_than_buffered() {
        let mut fits = BufReader::new(&b"0123456789\n"[..]);
        let Line::Text(line) = read_capped_line(&mut fits, 10).await else {
            panic!("a line that ends exactly at the cap is a line");
        };
        assert_eq!(line, "0123456789");

        let over = format!("{}\n", "x".repeat(11));
        let mut over = BufReader::new(over.as_bytes());
        assert!(matches!(read_capped_line(&mut over, 10).await, Line::Refused));

        // What is left of the stream is the proof: the cap stopped the read
        // rather than the end of the bytes.
        let endless = "x".repeat(64 * 1024);
        let mut source: &[u8] = endless.as_bytes();
        let mut reader = BufReader::with_capacity(64, &mut source);
        assert!(matches!(
            read_capped_line(&mut reader, 10).await,
            Line::Refused
        ));
        drop(reader);
        assert!(source.len() > endless.len() - 1024);
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

    /// A key that cannot be read is not a key that is gone. Minting over one
    /// would hand every peer holding this install as a contact an EndpointId
    /// that no longer answers, and would do it over a mode bit.
    #[cfg(unix)]
    #[test]
    fn an_unreadable_key_file_is_left_exactly_as_it_is() {
        use std::os::unix::fs::PermissionsExt;
        let path = scratch("unreadable");
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, KNOWN).unwrap();
        std::fs::set_permissions(&path, std::fs::Permissions::from_mode(0o000)).unwrap();
        assert!(
            std::fs::read(&path).is_err(),
            "this proves nothing unless the file is unreadable"
        );

        assert!(secret_key_at(&path).is_none());

        std::fs::set_permissions(&path, std::fs::Permissions::from_mode(0o600)).unwrap();
        assert_eq!(std::fs::read_to_string(&path).unwrap(), KNOWN);
    }

    /// The key lands by rename, so a write cut short leaves the old key in
    /// place rather than half of a new one.
    #[test]
    fn a_minted_key_leaves_no_temp_file_behind() {
        let path = scratch("atomic");
        secret_key_at(&path).expect("a minted key");

        let left: Vec<_> = std::fs::read_dir(path.parent().unwrap())
            .unwrap()
            .map(|entry| entry.unwrap().file_name())
            .collect();
        assert_eq!(left, vec![std::ffi::OsString::from("identity.key")]);
    }

    /// The replacement lands by rename, so a write that cannot get that far
    /// leaves what is on disk exactly as it was. A write that truncates first
    /// spends the stored key before it knows it can replace it.
    #[test]
    fn a_write_that_cannot_land_leaves_the_stored_bytes_alone() {
        let path = scratch("blocked-write");
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, "not a key").unwrap();
        // A directory where the temp file goes: the target stays writable, so
        // only the leg that is supposed to come first can fail.
        std::fs::create_dir_all(path.parent().unwrap().join(".identity.key.tmp")).unwrap();

        assert!(secret_key_at(&path).is_none());
        assert_eq!(std::fs::read_to_string(&path).unwrap(), "not a key");
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

    /// What the pump would have put on the webview's event bus.
    #[derive(Default)]
    struct Recorder {
        seen: Mutex<Vec<Event>>,
    }

    impl Events for Recorder {
        fn emit(&self, event: Event) {
            self.seen.lock().push(event);
        }
    }

    impl Recorder {
        fn peers(&self) -> Vec<String> {
            self.pick(|event| match event {
                Event::Peer(e) => Some(e.conn_id.clone()),
                _ => None,
            })
        }

        fn messages(&self) -> Vec<String> {
            self.pick(|event| match event {
                Event::Message(e) => Some(e.payload.clone()),
                _ => None,
            })
        }

        fn closed(&self) -> Vec<String> {
            self.pick(|event| match event {
                Event::Closed(e) => Some(e.conn_id.clone()),
                _ => None,
            })
        }

        fn pick(&self, of: impl Fn(&Event) -> Option<String>) -> Vec<String> {
            self.seen.lock().iter().filter_map(of).collect()
        }

        /// The pump runs on the endpoint's own runtime, so a test waits on it
        /// rather than driving it.
        fn wait(&self, what: &str, done: impl Fn(&Self) -> bool) {
            let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
            while std::time::Instant::now() < deadline {
                if done(self) {
                    return;
                }
                std::thread::sleep(std::time::Duration::from_millis(5));
            }
            panic!("nothing delivered {what} in ten seconds");
        }
    }

    /// The peer, on a runtime of its own so the endpoint under test is driven
    /// by nothing but its own accept loop.
    struct Guest {
        runtime: tokio::runtime::Runtime,
        endpoint: Endpoint,
    }

    impl Guest {
        fn new() -> Self {
            let runtime = tokio::runtime::Runtime::new().expect("runtime");
            let endpoint = runtime.block_on(bound());
            Self { runtime, endpoint }
        }

        fn dial(
            &self,
            to: EndpointAddr,
        ) -> (
            Connection,
            iroh::endpoint::SendStream,
            iroh::endpoint::RecvStream,
        ) {
            self.runtime.block_on(async {
                let conn = self.endpoint.connect(to, ALPN).await.expect("connect");
                let (send, recv) = conn.open_bi().await.expect("open_bi");
                (conn, send, recv)
            })
        }

        fn write(&self, send: &mut iroh::endpoint::SendStream, bytes: &[u8]) {
            let _ = self.runtime.block_on(async { send.write_all(bytes).await });
        }
    }

    fn live_addr(state: &CollabState) -> EndpointAddr {
        state
            .live
            .lock()
            .as_ref()
            .expect("a live endpoint")
            .endpoint
            .addr()
    }

    fn holds(state: &CollabState, conn_id: &str) -> bool {
        state
            .live
            .lock()
            .as_ref()
            .is_some_and(|live| live.conns.lock().contains_key(conn_id))
    }

    /// A connection the host let go of stops being a connection on the peer's
    /// side too, which is the whole difference between dropping a sender and
    /// closing a link.
    fn wait_closed(conn: &Connection) {
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
        while std::time::Instant::now() < deadline {
            if conn.close_reason().is_some() {
                return;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
        panic!("the peer is still holding an open connection");
    }

    /// The reader starts before the handshake has admitted anyone, so an
    /// unbounded line is reachable by anything that can open the ALPN.
    #[test]
    fn a_line_past_the_cap_drops_the_connection() {
        let state = CollabState::default();
        let events = Arc::new(Recorder::default());
        start(&state, events.clone(), false, false).expect("bind");

        let guest = Guest::new();
        let (conn, mut send, _recv) = guest.dial(live_addr(&state));
        // `accept_bi` returns only once the peer writes, so the flood both
        // opens the stream and overruns it.
        guest.write(&mut send, &vec![b'x'; MAX_LINE + 1]);

        events.wait("the close", |seen| !seen.closed().is_empty());
        let conn_id = events.peers()[0].clone();
        assert_eq!(events.closed(), vec![conn_id.clone()]);
        assert!(events.messages().is_empty(), "nothing past the cap parses");
        assert!(!holds(&state, &conn_id));
        wait_closed(&conn);

        stop(&state).expect("stop");
    }

    /// Dropping the sender leaves the peer connected and still being read, so
    /// a host that refused someone would keep hearing from them.
    #[test]
    fn closing_a_connection_takes_the_connection_with_it() {
        let state = CollabState::default();
        let events = Arc::new(Recorder::default());
        start(&state, events.clone(), false, false).expect("bind");

        let guest = Guest::new();
        let (conn, mut send, _recv) = guest.dial(live_addr(&state));
        guest.write(&mut send, b"{\"type\":\"hello\",\"protocol\":1}\n");
        events.wait("the hello", |seen| !seen.messages().is_empty());
        let conn_id = events.peers()[0].clone();
        assert!(holds(&state, &conn_id));

        close(&state, &conn_id).expect("close");

        assert!(!holds(&state, &conn_id));
        wait_closed(&conn);

        stop(&state).expect("stop");
    }

    /// Two PeerLinks share one bind, so a stop by one of them is a release and
    /// not a teardown.
    #[test]
    fn one_holder_stopping_leaves_the_endpoint_usable() {
        let state = CollabState::default();
        let events = Arc::new(Recorder::default());
        let listener = start(&state, events.clone(), false, false).expect("bind");
        let session = start(&state, events.clone(), false, false).expect("share");
        assert_eq!(listener, session, "one bind, not two");

        stop(&state).expect("the listener lets go");

        let guest = Guest::new();
        let (_conn, mut send, _recv) = guest.dial(live_addr(&state));
        guest.write(&mut send, b"{\"type\":\"hello\",\"protocol\":1}\n");
        events.wait("the hello", |seen| !seen.messages().is_empty());

        stop(&state).expect("the session lets go");
        assert!(state.live.lock().is_none(), "the last one out closes it");
    }
}
