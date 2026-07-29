#![forbid(unsafe_code)]

mod projection;
mod semantic;
#[cfg(all(target_arch = "wasm32", feature = "wasm"))]
pub mod wasm;

pub use projection::{
    AttributedComment, AttributedRevision, BookmarkFact, CommentContent, DocumentPackageProjection,
    DocumentParts, DocumentProjection, DocumentReviewFacts, DocumentStructureFacts, DocxLimits,
    FormattingProjectionStatus, FormattingUnknownReason, InternalParagraphId,
    InternalReferenceFact, InternalReferenceRole, NumberingHierarchyFact, PackageParagraphId,
    ParagraphIdentityFacts, ParagraphIndentation, ParagraphIndentationFact,
    ParagraphOutlineLevelFact, ParagraphStructure, ProjectedParagraph, ProjectionError,
    ProjectionOptions, ReviewDetail, ReviewFactLimits, ReviewFactSet, ReviewFactUnknownReason,
    ReviewPoint, ReviewSpan, RevisionContent, RevisionFactKind, RevisionProjectionStatus,
    RevisionUnsupportedReason, RevisionView, SpanCoverage, StructuralFactSet,
    StructuralFactUnknownReason, StructuralSpan, TextFormattingSpan, TextMaterialization,
    TextStyle, extract_document_parts, extract_document_xml, is_semantic_highlight_color,
    project_document_xml, project_document_xml_with_options, project_docx,
    project_docx_with_options, project_docx_with_review_facts,
};
pub use semantic::{
    BlockLocation, InlineContext, PartCoverage, PartScan, RevisionKind, ScanError, ScanLimits,
    ScanWork, SegmentSource, TextBlock, TextSegment, scan_wordprocessing_part,
    scan_wordprocessing_part_with,
};
