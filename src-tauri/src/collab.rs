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
use std::str::FromStr;
use std::sync::Arc;

use iroh::endpoint::{presets, Connection, TransportAddrUsage};
use iroh::{Endpoint, EndpointAddr, EndpointId, RelayMode};
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
        let endpoint = Endpoint::builder(presets::Minimal)
            .alpns(vec![ALPN.to_vec()])
            .relay_mode(relay_mode(relay))
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
) -> Result<String, String> {
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
    let _ = app.emit(
        "collab:peer",
        PeerEvent {
            conn_id: conn_id.clone(),
            endpoint_id: endpoint_id.clone(),
            connection_type: kind,
        },
    );
    let conns = live.conns.clone();
    let handle = app.clone();
    let id = conn_id.clone();
    live.runtime
        .spawn(async move { spawn_conn(handle, conns, id, conn, send, recv) });
    Ok(conn_id)
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
