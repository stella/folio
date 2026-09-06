//! The single entry point across the WebAssembly boundary.
//!
//! One call shapes one run. The result is a flat `Int32Array` rather than a
//! structure per glyph because a page of Arabic is tens of thousands of glyphs
//! and an object each would cost more in boundary crossings than the shaping
//! itself. The layout is fixed and documented on the TypeScript side, which is
//! the only place that reads it.

use crate::{Direction, ShapeRequest, ShapingError, shape_run};
use wasm_bindgen::prelude::*;

/// Fields per glyph in the returned buffer.
const FIELDS_PER_GLYPH: usize = 6;

fn describe(error: ShapingError) -> JsValue {
    JsValue::from_str(match error {
        ShapingError::UnreadableFace => "the bytes are not a font this shaper can read",
        ShapingError::InvalidTag => "an OpenType feature tag was not four characters",
    })
}

/// Shape one run of text with one face.
///
/// Returns `[unitsPerEm, glyphCount, (glyphId, cluster, xAdvance, yAdvance,
/// xOffset, yOffset) * glyphCount]` in font units.
///
/// # Errors
///
/// A string describing why the run could not be shaped: bytes that are not a
/// readable face, or a feature tag that is not four characters. The caller gets
/// a rejection rather than a run of no glyphs, which would paint as nothing.
// Every argument crosses the boundary as a scalar or a flat array. A request
// object would cost a call per field, and the whole reason this entry is shaped
// the way it is is to keep one crossing per run. Vectors are owned because
// wasm-bindgen materialises a `Vec<String>` from a JS string array; the body
// only borrows them.
#[allow(
    clippy::too_many_arguments,
    clippy::needless_pass_by_value,
    reason = "the WebAssembly boundary takes flat, owned arguments"
)]
#[wasm_bindgen(js_name = shapeRun)]
pub fn shape_run_wasm(
    font: &[u8],
    face_index: u32,
    text: &str,
    right_to_left: bool,
    script: &str,
    language: &str,
    features_on: Vec<String>,
    features_off: Vec<String>,
) -> Result<Vec<i32>, JsValue> {
    let request = ShapeRequest {
        font,
        face_index,
        text,
        direction: if right_to_left {
            Direction::RightToLeft
        } else {
            Direction::LeftToRight
        },
        script,
        language,
        features_on: &features_on,
        features_off: &features_off,
    };
    let shaped = shape_run(&request).map_err(describe)?;

    let mut out = Vec::with_capacity(
        shaped
            .glyphs
            .len()
            .saturating_mul(FIELDS_PER_GLYPH)
            .saturating_add(2),
    );
    out.push(shaped.units_per_em);
    out.push(i32::try_from(shaped.glyphs.len()).unwrap_or(i32::MAX));
    for glyph in &shaped.glyphs {
        out.push(i32::try_from(glyph.glyph_id).unwrap_or(0));
        out.push(i32::try_from(glyph.cluster).unwrap_or(0));
        out.push(glyph.x_advance);
        out.push(glyph.y_advance);
        out.push(glyph.x_offset);
        out.push(glyph.y_offset);
    }
    Ok(out)
}
