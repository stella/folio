//! The Unicode Bidirectional Algorithm (UAX #9) over one line.
//!
//! A shaped run must not span a direction change, so whoever splits text into
//! runs needs the embedding level of every character first, and whoever paints
//! the runs needs their visual order. Both answers come from here, for text
//! that is one paragraph laid out as one line: rules P2 and P3 (when the base
//! direction is automatic) through L2, with isolates, embeddings and overrides.
//! Splitting paragraphs and breaking lines stay with the caller.

use unicode_bidi::{Level, ParagraphBidiInfo};

/// The paragraph direction a caller asks for.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BaseDirection {
    LeftToRight,
    RightToLeft,
    /// Rules P2 and P3: the first strong character outside an isolate decides,
    /// and text with none is left to right.
    Auto,
}

/// The resolved line: one level and one visual position per character.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BidiLine {
    /// The paragraph embedding level: 0 left to right, 1 right to left.
    pub paragraph_level: u8,
    /// Embedding level per character (Unicode scalar value), after rule L1.
    /// Characters rule X9 removes (explicit formatting characters and boundary
    /// neutrals) carry the level of their neighbourhood and draw nothing.
    pub levels: Vec<u8>,
    /// Character indices in visual order, left to right (rule L2).
    pub visual_order: Vec<usize>,
}

/// Resolve the levels and the visual order of one line of text.
///
/// The text is treated as a single paragraph: a paragraph separator inside it
/// does not start a new one.
#[must_use]
pub fn resolve_bidi(text: &str, base: BaseDirection) -> BidiLine {
    let default_level = match base {
        BaseDirection::LeftToRight => Some(Level::ltr()),
        BaseDirection::RightToLeft => Some(Level::rtl()),
        BaseDirection::Auto => None,
    };
    let info = ParagraphBidiInfo::new(text, default_level);
    // An empty text has no line to reorder.
    if text.is_empty() {
        return BidiLine {
            paragraph_level: info.paragraph_level.number(),
            levels: Vec::new(),
            visual_order: Vec::new(),
        };
    }
    let levels = info.reordered_levels_per_char(0..text.len());
    BidiLine {
        paragraph_level: info.paragraph_level.number(),
        visual_order: ParagraphBidiInfo::reorder_visual(&levels),
        levels: levels.iter().map(Level::number).collect(),
    }
}
