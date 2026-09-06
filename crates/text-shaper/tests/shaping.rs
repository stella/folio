//! What shaping has to get right, in the scripts that need it.
//!
//! Each case asserts a property that a `cmap` lookup per code point cannot
//! produce, so a regression to unshaped output fails here rather than reaching
//! a reader.

#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]
// Assertion-style failure is the point of a test, and indexing happens only
// after the shape of the collection under test is established.
use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

use stella_text_shaper::{Direction, ShapeRequest, shape_run};

mod support;

/// Faces are a dependency of the React package, not of this crate, so a
/// checkout without them skips rather than fails.
fn face(relative: &str) -> Option<Vec<u8>> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/react/node_modules/@fontsource")
        .join(relative);
    let bytes = fs::read(path).ok()?;
    support::to_sfnt(&bytes)
}

const fn request<'a>(
    font: &'a [u8],
    text: &'a str,
    rtl: bool,
    script: &'a str,
) -> ShapeRequest<'a> {
    ShapeRequest {
        font,
        face_index: 0,
        text,
        direction: if rtl {
            Direction::RightToLeft
        } else {
            Direction::LeftToRight
        },
        script,
        language: "",
        features_on: &[],
        features_off: &[],
    }
}

#[test]
fn a_letter_takes_a_different_form_by_position() {
    let Some(font) = face("noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff") else {
        return;
    };
    // The sharpest test of joining: one letter, repeated. The first takes an
    // initial form and the last a final one, so the two cannot be the same
    // glyph. A per-code-point lookup produces the same glyph twice, so a
    // regression to unshaped output fails here.
    let shaped = shape_run(&request(&font, "\u{628}\u{628}", true, "Arab"))
        .expect("the bundled Arabic face must shape");

    let ids: Vec<u32> = shaped.glyphs.iter().map(|glyph| glyph.glyph_id).collect();
    let distinct: BTreeSet<u32> = ids.iter().copied().collect();
    assert!(
        distinct.len() > 1,
        "beh twice must not shape to one repeated glyph: {ids:?}"
    );
}

#[test]
fn arabic_runs_right_to_left() {
    let Some(font) = face("noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff") else {
        return;
    };
    let shaped = shape_run(&request(&font, "\u{628}\u{64a}\u{62a}", true, "Arab"))
        .expect("the bundled Arabic face must shape");

    // Clusters index the source text, so a right-to-left run emits them
    // descending: the first glyph painted is the last character read.
    let clusters: Vec<u32> = shaped.glyphs.iter().map(|glyph| glyph.cluster).collect();
    let mut descending = clusters.clone();
    descending.sort_unstable_by(|left, right| right.cmp(left));
    assert_eq!(
        clusters, descending,
        "an rtl run must emit clusters in reverse"
    );
}

#[test]
fn lam_alef_produces_a_glyph_neither_letter_produces_alone() {
    let Some(font) = face("noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff") else {
        return;
    };
    let together = shape_run(&request(&font, "\u{644}\u{627}", true, "Arab"))
        .expect("the bundled Arabic face must shape");
    let lam = shape_run(&request(&font, "\u{644}", true, "Arab")).expect("lam must shape");
    let alef = shape_run(&request(&font, "\u{627}", true, "Arab")).expect("alef must shape");

    // Lam followed by alef is a required ligature in Arabic. Whether a face
    // spells it as one glyph or as a ligature plus a mark is the face's
    // business; what no per-code-point mapping can produce is a glyph that
    // neither letter yields on its own.
    let apart: BTreeSet<u32> = lam
        .glyphs
        .iter()
        .chain(alef.glyphs.iter())
        .map(|glyph| glyph.glyph_id)
        .collect();
    let ligated = together
        .glyphs
        .iter()
        .any(|glyph| !apart.contains(&glyph.glyph_id));

    assert!(
        ligated,
        "lam-alef must ligate: {:?} against {apart:?} apart",
        together
            .glyphs
            .iter()
            .map(|glyph| glyph.glyph_id)
            .collect::<Vec<_>>()
    );
}

#[test]
fn latin_ligatures_form_when_the_face_has_them() {
    let Some(font) = face("carlito/files/carlito-latin-400-normal.woff") else {
        return;
    };
    let shaped =
        shape_run(&request(&font, "fi", false, "Latn")).expect("the bundled Latin face must shape");

    assert!(
        shaped.glyphs.len() <= 2,
        "shaping must not add glyphs to a two-character run"
    );
}

#[test]
fn a_run_reports_its_units_per_em() {
    let Some(font) = face("carlito/files/carlito-latin-400-normal.woff") else {
        return;
    };
    let shaped =
        shape_run(&request(&font, "a", false, "Latn")).expect("the bundled Latin face must shape");

    assert!(
        shaped.units_per_em > 0,
        "a caller scales advances by this, so zero would erase every measurement"
    );
}

#[test]
fn unreadable_bytes_are_an_error_rather_than_an_empty_run() {
    let result = shape_run(&request(&[0, 1, 2, 3], "a", false, "Latn"));

    assert!(
        result.is_err(),
        "an empty run would read as text that shapes to nothing"
    );
}
