//! The shaping call itself.
//!
//! Output is in font units, not pixels: a caller scales by its own font size,
//! so one shaped result serves a measurement in CSS pixels and a PDF text
//! matrix in points without either rounding the other's number.

use rustybuzz::{Direction as BuzzDirection, Face, Feature, Language, Script, UnicodeBuffer};

/// Which way the run advances. Resolved by the caller, never guessed here: the
/// producer already split runs at every direction change.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Direction {
    LeftToRight,
    RightToLeft,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShapingError {
    /// The bytes are not a font this shaper can read.
    UnreadableFace,
    /// A four-character OpenType tag was not four characters.
    InvalidTag,
}

/// One shaped glyph, in font units.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ShapedGlyph {
    /// Glyph id in the face that was shaped, ready for a `cmap`-free encoding.
    pub glyph_id: u32,
    /// First UTF-8 byte of the source text this glyph came from. A cluster of
    /// several glyphs shares one, which is what lets a caller map a glyph back
    /// to a character for text extraction.
    pub cluster: u32,
    pub x_advance: i32,
    pub y_advance: i32,
    pub x_offset: i32,
    pub y_offset: i32,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ShapedRun {
    pub glyphs: Vec<ShapedGlyph>,
    /// Font design units per em, so a caller can scale without re-reading the
    /// face it just handed in.
    pub units_per_em: i32,
}

pub struct ShapeRequest<'a> {
    pub font: &'a [u8],
    /// Index within a font collection; 0 for a plain font file.
    pub face_index: u32,
    pub text: &'a str,
    pub direction: Direction,
    /// ISO-15924 script tag, or empty to let the shaper detect it.
    pub script: &'a str,
    /// BCP-47 language, or empty. Affects locale-specific substitutions.
    pub language: &'a str,
    /// OpenType feature tags to force on, such as `liga` or `kern`.
    pub features_on: &'a [String],
    /// Feature tags to force off.
    pub features_off: &'a [String],
}

fn tag_of(value: &str) -> Option<rustybuzz::ttf_parser::Tag> {
    let bytes = value.as_bytes();
    let [a, b, c, d] = bytes else { return None };
    Some(rustybuzz::ttf_parser::Tag::from_bytes(&[*a, *b, *c, *d]))
}

fn features(request: &ShapeRequest<'_>) -> Result<Vec<Feature>, ShapingError> {
    let mut out = Vec::with_capacity(
        request
            .features_on
            .len()
            .saturating_add(request.features_off.len()),
    );
    for (tags, value) in [(request.features_on, 1_u32), (request.features_off, 0_u32)] {
        for tag in tags {
            let parsed = tag_of(tag).ok_or(ShapingError::InvalidTag)?;
            out.push(Feature::new(parsed, value, ..));
        }
    }
    Ok(out)
}

/// Shape one run. The run must not span a direction change: the producer
/// already splits at every boundary, so accepting one here would let a caller
/// hand over text whose visual order this cannot express.
///
/// # Errors
///
/// [`ShapingError::UnreadableFace`] when the bytes are not a face this can
/// read, [`ShapingError::InvalidTag`] when a requested feature tag is not four
/// characters. Both are refusals rather than an empty run, which a caller would
/// otherwise paint as text that shaped to nothing.
pub fn shape_run(request: &ShapeRequest<'_>) -> Result<ShapedRun, ShapingError> {
    let face =
        Face::from_slice(request.font, request.face_index).ok_or(ShapingError::UnreadableFace)?;
    let units_per_em = face.units_per_em();

    let mut buffer = UnicodeBuffer::new();
    buffer.push_str(request.text);
    if let Some(script) = tag_of(request.script).and_then(Script::from_iso15924_tag) {
        buffer.set_script(script);
    }
    if !request.language.is_empty()
        && let Ok(language) = request.language.parse::<Language>()
    {
        buffer.set_language(language);
    }
    // Fills in whatever the caller left unset. Without it the buffer keeps a
    // script of `Common`, and the shaper then applies none of the per-script
    // work that is the whole point of shaping: Arabic letters keep their
    // isolated forms, Indic syllables are not reordered.
    buffer.guess_segment_properties();
    // Direction last, so the caller's own answer wins over the guess: the
    // producer already split runs at every direction change and knows which
    // way this one runs.
    buffer.set_direction(match request.direction {
        Direction::LeftToRight => BuzzDirection::LeftToRight,
        Direction::RightToLeft => BuzzDirection::RightToLeft,
    });

    let shaped = rustybuzz::shape(&face, &features(request)?, buffer);
    let infos = shaped.glyph_infos();
    let positions = shaped.glyph_positions();
    let mut glyphs = Vec::with_capacity(infos.len());
    for (info, position) in infos.iter().zip(positions.iter()) {
        glyphs.push(ShapedGlyph {
            glyph_id: info.glyph_id,
            cluster: info.cluster,
            x_advance: position.x_advance,
            y_advance: position.y_advance,
            x_offset: position.x_offset,
            y_offset: position.y_offset,
        });
    }
    Ok(ShapedRun {
        glyphs,
        units_per_em,
    })
}
