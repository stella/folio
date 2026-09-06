#![forbid(unsafe_code)]
//! Complex-script text shaping.
//!
//! One shaper serves both consumers that need to know where a glyph goes: the
//! measurement seam that decides where lines break, and the PDF backend that
//! paints them. Two shapers would let a document paginate against one set of
//! advances and be painted with another, which is the divergence the whole
//! paint pipeline is arranged to prevent.
//!
//! Shaping is what turns a run of characters into glyphs and positions.
//! Latin needs it for ligatures and kerning; Arabic, Syriac and N'Ko select a
//! letter's initial, medial, final or isolated form from its neighbours;
//! Devanagari reorders matras and forms conjuncts; Thai stacks marks; Hebrew
//! positions points. None of those can be derived from a code point alone, at
//! any level of font-table reading, which is why this exists rather than a
//! `cmap` lookup.

mod shape;

#[cfg(all(target_arch = "wasm32", feature = "wasm"))]
pub mod wasm;

pub use shape::{Direction, ShapeRequest, ShapedGlyph, ShapedRun, ShapingError, shape_run};
