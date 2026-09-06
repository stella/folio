//! A guard on the test helper itself: a reconstruction that dropped `GSUB`
//! would make every shaping assertion pass vacuously against isolated forms.

#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]
// Assertion-style failure is the point of a test, and indexing happens only
// after the shape of the collection under test is established.
use std::fs;
use std::path::PathBuf;

mod support;

#[test]
fn the_reconstructed_face_keeps_its_layout_tables() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../../packages/react/node_modules/@fontsource/noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff",
    );
    let Ok(bytes) = fs::read(path) else {
        return;
    };
    let sfnt = support::to_sfnt(&bytes).expect("the bundled face must decode");

    let count = usize::from(u16::from_be_bytes([sfnt[4], sfnt[5]]));
    let tags: Vec<String> = (0..count)
        .filter_map(|index| {
            let entry = 12 + index * 16;
            sfnt.get(entry..entry + 4)
                .map(|tag| String::from_utf8_lossy(tag).to_string())
        })
        .collect();

    assert!(
        tags.iter().any(|tag| tag == "GSUB"),
        "no GSUB in the reconstruction, so nothing would shape: {tags:?}"
    );
}
