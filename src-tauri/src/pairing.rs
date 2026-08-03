//! A short code that is an address.
//!
//! Two debaters cannot exchange a 275-character ticket across a room, and most
//! of them have no application in common to paste one through. A code is eight
//! characters they can read aloud, and it works because both sides can compute
//! the same things from it: one HKDF gives a key, whose public half is the
//! EndpointId the host binds a temporary endpoint with, and a second HKDF over
//! a different label picks the relay that endpoint homes on.
//!
//! The relay is derived rather than defaulted because a default homes each node
//! on whichever relay is nearest. A host in one region and a guest in another
//! would pin different ones and never meet, which is precisely the failure this
//! exists to repair. Deriving it gives the two sides the same answer at any
//! distance.
//!
//! Nothing is published and nothing is looked up. The code travels by hand,
//! exactly as a ticket does, so this is addressing a person carries and not a
//! registry.
//!
//! This is the only implementation. TypeScript sends the code string and gets
//! an endpoint to dial: two implementations of one derivation must agree byte
//! for byte and they drift, and the web build cannot bind an endpoint at all.

use std::str::FromStr;

use hkdf::Hkdf;
use iroh::{RelayUrl, SecretKey};
use sha2::Sha256;

/// Eight characters of a 32-symbol alphabet, which is 40 bits. Guessing one
/// costs a QUIC dial per attempt against a code that lives ten minutes.
pub const CODE_LEN: usize = 8;

/// Crockford base32: no I, L, O or U. The first three are what a debater
/// reading a code aloud would have to disambiguate from 1 and 0, and the
/// fourth is left out so a code cannot spell an unfortunate word.
const ALPHABET: &[u8; 32] = b"0123456789ABCDEFGHJKMNPQRSTVWXYZ";

/// Names this derivation. A later scheme takes a new salt, so a code minted by
/// one build can never be misread by the other.
const SALT: &[u8] = b"ebb-pair-v1";

/// The relays a code can name.
///
/// Exactly iroh's default production map, so both sides can always reach the
/// one a code picks: a host homes there with `RelayMode::Custom`, and a guest
/// on `RelayMode::Default` already holds all four. A server of ebb's own would
/// be a backend, which this application does not have.
pub const RELAYS: [&str; 4] = [
    "https://use1-1.relay.n0.iroh.link./",
    "https://usw1-1.relay.n0.iroh.link./",
    "https://euc1-1.relay.n0.iroh.link./",
    "https://aps1-1.relay.n0.iroh.link./",
];

/// A fresh code. `% 32` is uniform because 32 divides 256 exactly, so no draw
/// is worth more than any other and the alphabet needs no rejection sampling.
pub fn new_code() -> String {
    rand::random::<[u8; CODE_LEN]>()
        .iter()
        .map(|b| ALPHABET[(*b % 32) as usize] as char)
        .collect()
}

/// The code behind whatever a debater typed: any case, and any spacing or
/// dashes the screen showed it with.
///
/// A character outside the alphabet is refused rather than mapped. Crockford
/// usually reads I and L as 1 and O as 0, but a code that silently becomes a
/// different code sends a debater to an endpoint nobody bound, and "that code
/// is wrong" is a better answer than a dial that times out.
pub fn normalize(raw: &str) -> Result<String, String> {
    let mut out = String::with_capacity(CODE_LEN);
    for ch in raw.chars() {
        if ch == '-' || ch.is_whitespace() {
            continue;
        }
        let up = ch.to_ascii_uppercase();
        if !up.is_ascii() || !ALPHABET.contains(&(up as u8)) {
            return Err("That code has a character ebb's codes never use".to_string());
        }
        if out.len() == CODE_LEN {
            return Err(format!("A code is {CODE_LEN} characters"));
        }
        out.push(up);
    }
    if out.len() != CODE_LEN {
        return Err(format!("A code is {CODE_LEN} characters"));
    }
    Ok(out)
}

/// One HKDF-SHA256 over the code, under one label.
fn seed(code: &str, info: &[u8], out: &mut [u8]) {
    Hkdf::<Sha256>::new(Some(SALT), code.as_bytes())
        .expand(info, out)
        .expect("a fixed output length well under HKDF's ceiling");
}

/// The key and the relay a code names.
pub fn derive(code: &str) -> Result<(SecretKey, RelayUrl), String> {
    derive_at(code, &RELAYS)
}

/// The same, against a caller's relay list, so a test can pin one that never
/// resolves rather than reaching for a real server.
pub fn derive_at(code: &str, relays: &[&str]) -> Result<(SecretKey, RelayUrl), String> {
    let code = normalize(code)?;
    if relays.is_empty() {
        return Err("No relay is configured for pairing".to_string());
    }
    let mut key_seed = [0u8; 32];
    seed(&code, b"key", &mut key_seed);
    // Its own label, so which relay a code lands on discloses nothing about the
    // key that code also names.
    let mut relay_seed = [0u8; 1];
    seed(&code, b"relay", &mut relay_seed);
    let pick = relays[relay_seed[0] as usize % relays.len()];
    let url = RelayUrl::from_str(pick).map_err(|_| "That relay is not a URL".to_string())?;
    Ok((SecretKey::from_bytes(&key_seed), url))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::collab::encode_hex;

    /// One code, one key, forever.
    ///
    /// Two installs on different releases must reach the same bytes from the
    /// same eight characters, so this is pinned rather than round-tripped: a
    /// refactor that changed the salt, the labels or the order would still
    /// round-trip against itself and would silently stop two debaters meeting.
    #[test]
    fn a_code_derives_one_key_for_good() {
        let (key, _) = derive("K7QM3XPV").expect("a valid code");
        assert_eq!(
            encode_hex(&key.to_bytes()),
            "32242047f09d1425c18b7041edeeb155dd68ad5c8bf084726faa04c986c76a28"
        );
    }

    #[test]
    fn a_code_reads_the_same_however_it_is_typed() {
        let plain = derive("K7QM3XPV").expect("a valid code").0.to_bytes();
        for typed in ["k7qm3xpv", "K7QM-3XPV", "k7qm 3xpv", " K7QM-3xpv "] {
            assert_eq!(
                derive(typed).expect("a valid code").0.to_bytes(),
                plain,
                "{typed}"
            );
        }
    }

    #[test]
    fn two_codes_do_not_collide() {
        let a = derive("K7QM3XPV").expect("a valid code").0.to_bytes();
        let b = derive("K7QM3XPW").expect("a valid code").0.to_bytes();
        assert_ne!(a, b);
    }

    #[test]
    fn the_relay_comes_from_a_different_label_than_the_key() {
        // Distinct labels, so which relay a code lands on says nothing about
        // its key. Pinned for the reason the key is.
        assert_eq!(
            derive("K7QM3XPV").expect("a valid code").1.to_string(),
            RELAYS[0]
        );
        assert_eq!(
            derive("ZZZZZZZZ").expect("a valid code").1.to_string(),
            RELAYS[1]
        );
    }

    #[test]
    fn every_derived_relay_is_one_this_build_ships() {
        for code in ["K7QM3XPV", "ZZZZZZZZ", "00000000", "ABCDEFGH", "TESTAA01"] {
            let (_, relay) = derive(code).expect("a valid code");
            assert!(
                RELAYS.contains(&relay.to_string().as_str()),
                "{code} landed off the list"
            );
        }
    }

    #[test]
    fn the_ambiguous_characters_are_refused() {
        // Crockford base32 leaves out I, L, O and U, so a debater reading a
        // code aloud never has to say which one they meant.
        for bad in ["IIIIIIII", "LLLLLLLL", "OOOOOOOO", "UUUUUUUU", "K7QM3XPI"] {
            assert!(normalize(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn a_code_is_eight_characters_and_not_seven_or_nine() {
        assert!(normalize("K7QM3XP").is_err());
        assert!(normalize("K7QM3XPVV").is_err());
        assert!(normalize("").is_err());
        assert_eq!(normalize("K7QM3XPV").expect("a valid code"), "K7QM3XPV");
    }

    #[test]
    fn nothing_but_the_alphabet_gets_in() {
        for bad in ["K7QM3XP!", "K7QM3XP\u{00e9}", "<script>", "K7QM_3XPV"] {
            assert!(normalize(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn a_minted_code_is_always_one_this_build_would_accept() {
        for _ in 0..2000 {
            let code = new_code();
            assert_eq!(code.len(), CODE_LEN);
            assert_eq!(normalize(&code).as_deref(), Ok(code.as_str()));
        }
    }

    #[test]
    fn a_minted_code_is_not_the_same_code_every_time() {
        let mut seen = std::collections::HashSet::new();
        for _ in 0..64 {
            seen.insert(new_code());
        }
        assert!(seen.len() > 60, "{} distinct out of 64", seen.len());
    }

    #[test]
    fn a_relay_list_that_is_empty_is_refused_rather_than_dividing_by_zero() {
        assert!(derive_at("K7QM3XPV", &[]).is_err());
    }

    #[test]
    fn a_relay_that_will_not_parse_is_refused() {
        assert!(derive_at("K7QM3XPV", &["not a url"]).is_err());
    }
}
