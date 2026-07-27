#![forbid(unsafe_code)]

mod projection;
mod semantic;
#[cfg(all(target_arch = "wasm32", feature = "wasm"))]
pub mod wasm;

pub use projection::{
    BookmarkFact, DocumentParts, DocumentProjection, DocumentStructureFacts, DocxLimits,
    InternalParagraphId, InternalReferenceFact, InternalReferenceRole, NumberingHierarchyFact,
    PackageParagraphId, ParagraphIdentityFacts, ParagraphIndentation, ParagraphIndentationFact,
    ParagraphStructure, ProjectedParagraph, ProjectionError, ProjectionOptions,
    RevisionProjectionStatus, RevisionUnsupportedReason, RevisionView, SpanCoverage,
    StructuralFactSet, StructuralFactUnknownReason, StructuralSpan, TextFormattingSpan,
    TextMaterialization, TextStyle, extract_document_parts, extract_document_xml,
    is_semantic_highlight_color, project_document_xml, project_document_xml_with_options,
    project_docx, project_docx_with_options,
};
pub use semantic::{
    BlockLocation, InlineContext, PartCoverage, PartScan, RevisionKind, ScanError, ScanLimits,
    ScanWork, SegmentSource, TextBlock, TextSegment, scan_wordprocessing_part,
    scan_wordprocessing_part_with,
};
