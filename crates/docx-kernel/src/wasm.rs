use crate::{
    AttributedComment, AttributedRevision, BookmarkFact, DocumentPackageProjection,
    DocumentProjection, DocumentReviewFacts, DocumentStructureFacts, DocxLimits,
    FormattingProjectionStatus, FormattingUnknownReason, InternalParagraphId,
    InternalReferenceFact, InternalReferenceRole, NumberingHierarchyFact, ParagraphAlignmentFact,
    ParagraphAlignmentSource, ParagraphAlignmentValue, ParagraphIdentityFacts,
    ParagraphIndentationFact, ParagraphOutlineLevelFact, ParagraphStructure, ProjectedParagraph,
    ProjectionOptions, ReviewDetail, ReviewFactLimits, ReviewFactSet, ReviewFactUnknownReason,
    ReviewSpan, RevisionFactKind, RevisionProjectionStatus, RevisionUnsupportedReason,
    SpanCoverage, StructuralFactSet, StructuralFactUnknownReason, StructuralSpan,
    TextMaterialization, TextStyle, project_docx, project_docx_with_review_facts,
};
use js_sys::Array;
use wasm_bindgen::{JsCast, prelude::*};

const DOCX_PROJECTION_SCHEMA_VERSION: u32 = 5;
const DOCX_PACKAGE_PROJECTION_SCHEMA_VERSION: u32 = 2;
const DOCX_REVIEW_FACTS_SCHEMA_VERSION: u32 = 2;

#[wasm_bindgen(typescript_custom_section)]
const TYPESCRIPT_TYPES: &str = r#"
export type DocxProjectionFormattingSpan = readonly [
  startUtf16: number,
  endUtf16: number,
  style: "bold" | "highlight" | "superscript",
];
export type DocxProjectionStructure =
  | readonly []
  | readonly [
      type: "table",
      tableId: string,
      row: number,
      column: number,
    ];
export type DocxProjectionAlignmentValue = "center" | "justify" | "left" | "right";
export type DocxProjectionAlignmentSource = "direct" | "style";
export type DocxProjectionAlignment = readonly [
  value: DocxProjectionAlignmentValue,
  source: DocxProjectionAlignmentSource,
] | null;
export type DocxProjectionParagraph = readonly [
  ordinal: number,
  text: string,
  packageParagraphId: string | null,
  formatting: readonly DocxProjectionFormattingSpan[],
  structure: DocxProjectionStructure,
  styleId: string | null,
  alignment: DocxProjectionAlignment,
];
export type DocxProjectionFormattingUnknownReason =
  | "document-part-only"
  | "styles-part-unavailable"
  | "unsupported-styles";
export type DocxProjectionFormattingStatus =
  | readonly [status: "complete"]
  | readonly [
      status: "incomplete",
      reason: DocxProjectionFormattingUnknownReason,
    ];
export type DocxProjectionFactSet<T> =
  | readonly [status: "known", items: readonly T[]]
  | readonly [status: "unknown", reason: DocxProjectionUnknownReason];
export type DocxProjectionUnknownReason =
  | "document-part-only"
  | "styles-part-unavailable"
  | "unsupported-styles"
  | "unsupported-numbering"
  | "incomplete-bookmark-ranges"
  | "unsupported-internal-references";
export type DocxProjectionIndentationFact = readonly [
  paragraphOrdinal: number,
  firstLineTwips: number | null,
  hangingTwips: number | null,
  leftTwips: number | null,
  rightTwips: number | null,
  startTwips: number | null,
  endTwips: number | null,
  firstLineCharsHundredths: number | null,
  hangingCharsHundredths: number | null,
  leftCharsHundredths: number | null,
  rightCharsHundredths: number | null,
  startCharsHundredths: number | null,
  endCharsHundredths: number | null,
];
export type DocxProjectionNumberingFact = readonly [
  paragraphOrdinal: number,
  parentParagraphOrdinal: number | null,
  childParagraphOrdinals: readonly number[],
];
export type DocxProjectionOutlineLevelFact = readonly [
  paragraphOrdinal: number,
  outlineLevel: number,
];
export type DocxProjectionBookmarkFact = readonly [
  paragraphOrdinal: number,
  bookmarkId: number,
  name: string,
  span: DocxProjectionStructuralSpan,
];
export type DocxProjectionReferenceFact = readonly [
  paragraphOrdinal: number,
  referenceId: string,
  role: "source" | "target",
  span: DocxProjectionStructuralSpan,
];
export type DocxProjectionStructuralSpan = readonly [
  startUtf8: number,
  endUtf8: number,
  startUtf16: number,
  endUtf16: number,
  coverage:
    | "complete"
    | "continues-before"
    | "continues-after"
    | "continues-before-and-after",
];
export type DocxProjectionStructuralFacts = readonly [
  indentation: DocxProjectionFactSet<DocxProjectionIndentationFact>,
  numberingHierarchy: DocxProjectionFactSet<DocxProjectionNumberingFact>,
  bookmarks: DocxProjectionFactSet<DocxProjectionBookmarkFact>,
  internalReferences: DocxProjectionFactSet<DocxProjectionReferenceFact>,
  outlineLevels: DocxProjectionFactSet<DocxProjectionOutlineLevelFact>,
];
export type DocxProjectionRevisionUnsupportedReason =
  | "incompatible-paragraph-merge"
  | "structural-table-revision"
  | "unsupported-revision-markup";
export type DocxProjectionRevisionStatus =
  | readonly [status: "complete"]
  | readonly [
      status: "incomplete",
      reasons: readonly DocxProjectionRevisionUnsupportedReason[],
    ];
export type DocxProjectionWire = readonly [
  schemaVersion: 5,
  paragraphs: readonly DocxProjectionParagraph[],
  structuralFacts: DocxProjectionStructuralFacts,
  revisionStatus: DocxProjectionRevisionStatus,
  formattingStatus: DocxProjectionFormattingStatus,
];
export type DocxReviewUnknownReason =
  | "invalid-document"
  | "invalid-comments"
  | "invalid-comments-extended"
  | "resource-limit"
  | "unsupported-location";
export type DocxReviewFactSet<T> =
  | readonly [status: "known", items: readonly T[]]
  | readonly [status: "unknown", reason: DocxReviewUnknownReason];
export type DocxRevisionKind =
    | "insertion"
    | "deletion"
    | "moveFrom"
    | "moveTo"
    | "cellIns"
    | "cellDel"
    | "cellMerge"
    | "pPrChange"
    | "rPrChange"
    | "sectPrChange"
    | "tblPrChange"
    | "trPrChange"
    | "tcPrChange"
    | "tblGridChange"
    | "customXmlDelRangeStart"
    | "customXmlDelRangeEnd"
    | "customXmlInsRangeStart"
    | "customXmlInsRangeEnd"
    | "customXmlMoveFromRangeStart"
    | "customXmlMoveFromRangeEnd"
    | "customXmlMoveToRangeStart"
    | "customXmlMoveToRangeEnd";
/**
 * Attributed revision wire tuple. Known content is flattened into the item so
 * one review fact creates one JavaScript array at the WASM boundary.
 */
export type DocxAttributedRevision =
  | readonly [
      type: DocxRevisionKind,
      author: string,
      date: string | null,
      revisionId: string | null,
      contentStatus: "known",
      startParagraphOrdinal: number,
      startUtf8: number,
      startUtf16: number,
      endParagraphOrdinal: number,
      endUtf8: number,
      endUtf16: number,
      text: string,
      contentKind: "text" | "formatting-only",
    ]
  | readonly [
      type: DocxRevisionKind,
      author: string,
      date: string | null,
      revisionId: string | null,
      contentStatus: "unknown",
      reason: DocxReviewUnknownReason,
    ];
/**
 * Attributed comment wire tuple. Known content is flattened into the item so
 * one review fact creates one JavaScript array at the WASM boundary.
 */
export type DocxAttributedComment =
  | readonly [
      commentId: string,
      author: string,
      initials: string | null,
      date: string | null,
      parentCommentId: string | null,
      threadState: "open" | "resolved",
      contentStatus: "known",
      startParagraphOrdinal: number,
      startUtf8: number,
      startUtf16: number,
      endParagraphOrdinal: number,
      endUtf8: number,
      endUtf16: number,
      commentText: string,
      referencedText: string,
    ]
  | readonly [
      commentId: string,
      author: string,
      initials: string | null,
      date: string | null,
      parentCommentId: string | null,
      threadState: "open" | "resolved",
      contentStatus: "unknown",
      reason: DocxReviewUnknownReason,
    ];
/** Review-fact wire schema. A new tuple layout requires a schema-version bump. */
export type DocxReviewFactsWire = readonly [
  schemaVersion: 2,
  revisions: DocxReviewFactSet<DocxAttributedRevision>,
  comments: DocxReviewFactSet<DocxAttributedComment>,
];
/** Fused package wire schema. A new tuple layout requires a schema-version bump. */
export type DocxPackageProjectionWire = readonly [
  schemaVersion: 2,
  document: DocxProjectionWire,
  reviewFacts: DocxReviewFactsWire,
];
"#;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(typescript_type = "DocxProjectionWire")]
    pub type DocxProjectionWire;
    #[wasm_bindgen(typescript_type = "DocxPackageProjectionWire")]
    pub type DocxPackageProjectionWire;
}

/// Projects compressed DOCX bytes into a versioned host-independent snapshot.
///
/// The ordinal is the paragraph's position in this immutable package snapshot.
/// Package `w14:paraId` values are returned separately. Neither value is an
/// application identity or a host navigation identity. A structural fact
/// family is `unknown` when the package does not prove completeness; an empty
/// `known` list is authoritative negative evidence.
///
/// # Errors
///
/// Returns a JavaScript `Error` when the package cannot be projected or a
/// numeric wire value cannot be represented by the schema.
#[wasm_bindgen(js_name = projectCompressedDocx)]
pub fn project_compressed_docx(bytes: &[u8]) -> Result<DocxProjectionWire, JsValue> {
    project_docx_projection(bytes)
        .and_then(|projection| output_projection_with_structure(&projection))
        // SAFETY: the output builder constructs the exact tuple declared as
        // `DocxProjectionWire` in the wasm-bindgen TypeScript custom section.
        .map(JsCast::unchecked_into)
        .map_err(|error| js_error(&error))
}

/// Projects the document snapshot and attributed review facts from one bounded
/// package-directory scan.
///
/// # Errors
///
/// Returns a JavaScript `Error` when the document package itself cannot be
/// projected. Invalid optional review parts are represented as unknown facts.
#[wasm_bindgen(js_name = projectCompressedDocxWithReviewFacts)]
pub fn project_compressed_docx_with_review_facts(
    bytes: &[u8],
) -> Result<DocxPackageProjectionWire, JsValue> {
    project_docx_package_projection(bytes, TextMaterialization::WordHost)
        .and_then(|projection| output_package_projection(&projection))
        // SAFETY: the output builder constructs the exact tuple declared as
        // `DocxPackageProjectionWire` in the TypeScript custom section.
        .map(JsCast::unchecked_into)
        .map_err(|error| js_error(&error))
}

/// Projects the same fused snapshot with controls normalized for readable text.
///
/// # Errors
///
/// Returns a JavaScript `Error` under the same conditions as
/// [`project_compressed_docx_with_review_facts`].
#[wasm_bindgen(js_name = projectCompressedDocxWithReadableReviewFacts)]
pub fn project_compressed_docx_with_readable_review_facts(
    bytes: &[u8],
) -> Result<DocxPackageProjectionWire, JsValue> {
    project_docx_package_projection(bytes, TextMaterialization::ReadablePlainText)
        .and_then(|projection| output_package_projection(&projection))
        // SAFETY: the output builder constructs the exact tuple declared as
        // `DocxPackageProjectionWire` in the TypeScript custom section.
        .map(JsCast::unchecked_into)
        .map_err(|error| js_error(&error))
}

fn project_docx_projection(bytes: &[u8]) -> Result<DocumentProjection, String> {
    let limits = DocxLimits::default();
    project_docx(bytes, limits, |facts: ParagraphIdentityFacts<'_>| {
        InternalParagraphId::new(format!("projected-{}", facts.ordinal))
    })
    .map_err(|error| error.to_string())
}

fn project_docx_package_projection(
    bytes: &[u8],
    text_materialization: TextMaterialization,
) -> Result<DocumentPackageProjection, String> {
    project_docx_with_review_facts(
        bytes,
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions {
            text_materialization,
            ..ProjectionOptions::default()
        },
        |facts: ParagraphIdentityFacts<'_>| {
            InternalParagraphId::new(format!("projected-{}", facts.ordinal))
        },
    )
    .map_err(|error| error.to_string())
}

fn output_paragraphs(projection: &DocumentProjection) -> Result<JsValue, String> {
    let output = Array::new();
    for (ordinal, paragraph) in projection.paragraphs.iter().enumerate() {
        output.push(&output_projected_paragraph(ordinal, paragraph)?);
    }
    Ok(output.into())
}

fn output_projected_paragraph(
    ordinal: usize,
    paragraph: &ProjectedParagraph,
) -> Result<JsValue, String> {
    let output = Array::new_with_length(7);
    output.set(0, usize_number(ordinal)?);
    output.set(1, JsValue::from_str(&paragraph.text));
    output.set(
        2,
        paragraph.package_paragraph_id.map_or(JsValue::NULL, |id| {
            JsValue::from_str(&format!("{:08X}", id.value()))
        }),
    );

    let formatting = Array::new();
    for span in &paragraph.formatting {
        let output_span = Array::new_with_length(3);
        output_span.set(0, JsValue::from_f64(f64::from(span.start_utf16)));
        output_span.set(1, JsValue::from_f64(f64::from(span.end_utf16)));
        output_span.set(2, JsValue::from_str(text_style_wire_name(span.style)));
        formatting.push(&output_span);
    }
    output.set(3, formatting.into());
    output.set(4, output_paragraph_structure(paragraph.structure.clone())?);
    output.set(
        5,
        paragraph
            .style_id
            .as_ref()
            .map_or(JsValue::NULL, |style_id| JsValue::from_str(style_id)),
    );
    output.set(6, output_paragraph_alignment(paragraph.alignment));
    Ok(output.into())
}

fn output_paragraph_alignment(alignment: Option<ParagraphAlignmentFact>) -> JsValue {
    alignment.map_or(JsValue::NULL, |alignment| {
        let output = Array::new_with_length(2);
        output.set(
            0,
            JsValue::from_str(alignment_value_wire_name(alignment.value)),
        );
        output.set(
            1,
            JsValue::from_str(alignment_source_wire_name(alignment.source)),
        );
        output.into()
    })
}

const fn alignment_value_wire_name(value: ParagraphAlignmentValue) -> &'static str {
    match value {
        ParagraphAlignmentValue::Center => "center",
        ParagraphAlignmentValue::Justify => "justify",
        ParagraphAlignmentValue::Left => "left",
        ParagraphAlignmentValue::Right => "right",
    }
}

const fn alignment_source_wire_name(source: ParagraphAlignmentSource) -> &'static str {
    match source {
        ParagraphAlignmentSource::Direct => "direct",
        ParagraphAlignmentSource::Style => "style",
    }
}

const fn text_style_wire_name(style: TextStyle) -> &'static str {
    match style {
        TextStyle::Bold => "bold",
        TextStyle::Highlight => "highlight",
        TextStyle::Superscript => "superscript",
    }
}

fn output_projection_with_structure(projection: &DocumentProjection) -> Result<JsValue, String> {
    let output = Array::new_with_length(5);
    output.set(
        0,
        JsValue::from_f64(f64::from(DOCX_PROJECTION_SCHEMA_VERSION)),
    );
    output.set(1, output_paragraphs(projection)?);
    output.set(2, output_structural_facts(&projection.structural_facts)?);
    output.set(3, output_revision_status(&projection.revision_status));
    output.set(4, output_formatting_status(projection.formatting_status));
    Ok(output.into())
}

fn output_formatting_status(status: FormattingProjectionStatus) -> JsValue {
    let output = Array::new();
    match status {
        FormattingProjectionStatus::Complete => {
            output.push(&JsValue::from_str("complete"));
        }
        FormattingProjectionStatus::Incomplete(reason) => {
            output.push(&JsValue::from_str("incomplete"));
            output.push(&JsValue::from_str(match reason {
                FormattingUnknownReason::DocumentPartOnly => "document-part-only",
                FormattingUnknownReason::StylesPartUnavailable => "styles-part-unavailable",
                FormattingUnknownReason::UnsupportedStyles => "unsupported-styles",
            }));
        }
    }
    output.into()
}

fn output_package_projection(projection: &DocumentPackageProjection) -> Result<JsValue, String> {
    let output = Array::new_with_length(3);
    output.set(
        0,
        JsValue::from_f64(f64::from(DOCX_PACKAGE_PROJECTION_SCHEMA_VERSION)),
    );
    output.set(1, output_projection_with_structure(&projection.document)?);
    output.set(2, output_review_facts(&projection.review_facts)?);
    Ok(output.into())
}

fn output_review_facts(facts: &DocumentReviewFacts) -> Result<JsValue, String> {
    let output = Array::new_with_length(3);
    output.set(
        0,
        JsValue::from_f64(f64::from(DOCX_REVIEW_FACTS_SCHEMA_VERSION)),
    );
    output.set(
        1,
        output_review_fact_set(&facts.revisions, output_revision_fact)?,
    );
    output.set(
        2,
        output_review_fact_set(&facts.comments, output_comment_fact)?,
    );
    Ok(output.into())
}

fn output_review_fact_set<T>(
    facts: &ReviewFactSet<T>,
    mut output_item: impl FnMut(&T) -> Result<JsValue, String>,
) -> Result<JsValue, String> {
    let output = Array::new_with_length(2);
    match facts {
        ReviewFactSet::Known(items) => {
            output.set(0, JsValue::from_str("known"));
            let values = Array::new();
            for item in items {
                values.push(&output_item(item)?);
            }
            output.set(1, values.into());
        }
        ReviewFactSet::Unknown(reason) => {
            output.set(0, JsValue::from_str("unknown"));
            output.set(1, JsValue::from_str(review_unknown_reason(*reason)));
        }
    }
    Ok(output.into())
}

fn output_revision_fact(fact: &AttributedRevision) -> Result<JsValue, String> {
    let output = Array::new_with_length(match &fact.content {
        ReviewDetail::Known(_) => 13,
        ReviewDetail::Unknown(_) => 6,
    });
    output.set(
        0,
        JsValue::from_str(match fact.kind {
            RevisionFactKind::Insertion => "insertion",
            RevisionFactKind::Deletion => "deletion",
            RevisionFactKind::MoveFrom => "moveFrom",
            RevisionFactKind::MoveTo => "moveTo",
            RevisionFactKind::CellInsertion => "cellIns",
            RevisionFactKind::CellDeletion => "cellDel",
            RevisionFactKind::CellMerge => "cellMerge",
            RevisionFactKind::ParagraphPropertiesChange => "pPrChange",
            RevisionFactKind::RunPropertiesChange => "rPrChange",
            RevisionFactKind::SectionPropertiesChange => "sectPrChange",
            RevisionFactKind::TablePropertiesChange => "tblPrChange",
            RevisionFactKind::TableRowPropertiesChange => "trPrChange",
            RevisionFactKind::TableCellPropertiesChange => "tcPrChange",
            RevisionFactKind::TableGridChange => "tblGridChange",
            RevisionFactKind::CustomXmlDeletionRangeStart => "customXmlDelRangeStart",
            RevisionFactKind::CustomXmlDeletionRangeEnd => "customXmlDelRangeEnd",
            RevisionFactKind::CustomXmlInsertionRangeStart => "customXmlInsRangeStart",
            RevisionFactKind::CustomXmlInsertionRangeEnd => "customXmlInsRangeEnd",
            RevisionFactKind::CustomXmlMoveFromRangeStart => "customXmlMoveFromRangeStart",
            RevisionFactKind::CustomXmlMoveFromRangeEnd => "customXmlMoveFromRangeEnd",
            RevisionFactKind::CustomXmlMoveToRangeStart => "customXmlMoveToRangeStart",
            RevisionFactKind::CustomXmlMoveToRangeEnd => "customXmlMoveToRangeEnd",
        }),
    );
    output.set(1, JsValue::from_str(&fact.author));
    output.set(2, optional_string(fact.date.as_deref()));
    output.set(3, optional_string(fact.revision_id.as_deref()));
    match &fact.content {
        ReviewDetail::Known(content) => {
            output.set(4, JsValue::from_str("known"));
            output_review_span_fields(&output, 5, content.span)?;
            output.set(11, JsValue::from_str(&content.text));
            output.set(
                12,
                JsValue::from_str(if content.formatting_only {
                    "formatting-only"
                } else {
                    "text"
                }),
            );
        }
        ReviewDetail::Unknown(reason) => {
            output.set(4, JsValue::from_str("unknown"));
            output.set(5, JsValue::from_str(review_unknown_reason(*reason)));
        }
    }
    Ok(output.into())
}

fn output_comment_fact(fact: &AttributedComment) -> Result<JsValue, String> {
    let output = Array::new_with_length(match &fact.content {
        ReviewDetail::Known(_) => 15,
        ReviewDetail::Unknown(_) => 8,
    });
    output.set(0, JsValue::from_str(&fact.comment_id));
    output.set(1, JsValue::from_str(&fact.author));
    output.set(2, optional_string(fact.initials.as_deref()));
    output.set(3, optional_string(fact.date.as_deref()));
    output.set(4, optional_string(fact.parent_comment_id.as_deref()));
    output.set(
        5,
        JsValue::from_str(if fact.resolved { "resolved" } else { "open" }),
    );
    match &fact.content {
        ReviewDetail::Known(content) => {
            output.set(6, JsValue::from_str("known"));
            output_review_span_fields(&output, 7, content.anchor)?;
            output.set(13, JsValue::from_str(&content.comment_text));
            output.set(14, JsValue::from_str(&content.referenced_text));
        }
        ReviewDetail::Unknown(reason) => {
            output.set(6, JsValue::from_str("unknown"));
            output.set(7, JsValue::from_str(review_unknown_reason(*reason)));
        }
    }
    Ok(output.into())
}

fn optional_string(value: Option<&str>) -> JsValue {
    value.map_or(JsValue::NULL, JsValue::from_str)
}

fn output_review_span_fields(
    output: &Array,
    start_index: u32,
    span: ReviewSpan,
) -> Result<(), String> {
    let fields = [
        usize_number(span.start.paragraph_ordinal)?,
        JsValue::from_f64(f64::from(span.start.utf8)),
        JsValue::from_f64(f64::from(span.start.utf16)),
        usize_number(span.end.paragraph_ordinal)?,
        JsValue::from_f64(f64::from(span.end.utf8)),
        JsValue::from_f64(f64::from(span.end.utf16)),
    ];
    for (offset, field) in fields.into_iter().enumerate() {
        let offset = u32::try_from(offset)
            .map_err(|_| "review span output index does not fit u32".to_owned())?;
        let index = start_index
            .checked_add(offset)
            .ok_or_else(|| "review span output index overflow".to_owned())?;
        output.set(index, field);
    }
    Ok(())
}

const fn review_unknown_reason(reason: ReviewFactUnknownReason) -> &'static str {
    match reason {
        ReviewFactUnknownReason::InvalidDocument => "invalid-document",
        ReviewFactUnknownReason::InvalidComments => "invalid-comments",
        ReviewFactUnknownReason::InvalidCommentsExtended => "invalid-comments-extended",
        ReviewFactUnknownReason::ResourceLimit => "resource-limit",
        ReviewFactUnknownReason::UnsupportedLocation => "unsupported-location",
    }
}

fn output_revision_status(status: &RevisionProjectionStatus) -> JsValue {
    let output = Array::new();
    match status {
        RevisionProjectionStatus::Complete => {
            output.push(&JsValue::from_str("complete"));
        }
        RevisionProjectionStatus::Incomplete(reasons) => {
            output.push(&JsValue::from_str("incomplete"));
            let output_reasons = Array::new();
            for reason in reasons {
                output_reasons.push(&JsValue::from_str(match reason {
                    RevisionUnsupportedReason::IncompatibleParagraphMerge => {
                        "incompatible-paragraph-merge"
                    }
                    RevisionUnsupportedReason::StructuralTableRevision => {
                        "structural-table-revision"
                    }
                    RevisionUnsupportedReason::UnsupportedRevisionMarkup => {
                        "unsupported-revision-markup"
                    }
                }));
            }
            output.push(&output_reasons);
        }
    }
    output.into()
}

fn output_structural_facts(facts: &DocumentStructureFacts) -> Result<JsValue, String> {
    let output = Array::new_with_length(5);
    output.set(
        0,
        output_fact_set(&facts.indentation, output_indentation_fact)?,
    );
    output.set(
        1,
        output_fact_set(&facts.numbering_hierarchy, output_numbering_fact)?,
    );
    output.set(2, output_fact_set(&facts.bookmarks, output_bookmark_fact)?);
    output.set(
        3,
        output_fact_set(&facts.internal_references, output_reference_fact)?,
    );
    output.set(
        4,
        output_fact_set(&facts.outline_levels, output_outline_level_fact)?,
    );
    Ok(output.into())
}

fn output_fact_set<T>(
    facts: &StructuralFactSet<T>,
    mut output_item: impl FnMut(&T) -> Result<JsValue, String>,
) -> Result<JsValue, String> {
    let output = Array::new_with_length(2);
    match facts {
        StructuralFactSet::Known(items) => {
            output.set(0, JsValue::from_str("known"));
            let output_items = Array::new();
            for item in items {
                output_items.push(&output_item(item)?);
            }
            output.set(1, output_items.into());
        }
        StructuralFactSet::Unknown(reason) => {
            output.set(0, JsValue::from_str("unknown"));
            output.set(1, JsValue::from_str(unknown_reason(*reason)));
        }
    }
    Ok(output.into())
}

fn output_indentation_fact(fact: &ParagraphIndentationFact) -> Result<JsValue, String> {
    let value = fact.value;
    let output = Array::new_with_length(13);
    output.set(0, usize_number(fact.paragraph_ordinal)?);
    for (item, index) in [
        value.first_line_twips,
        value.hanging_twips,
        value.left_twips,
        value.right_twips,
        value.start_twips,
        value.end_twips,
        value.first_line_chars_hundredths,
        value.hanging_chars_hundredths,
        value.left_chars_hundredths,
        value.right_chars_hundredths,
        value.start_chars_hundredths,
        value.end_chars_hundredths,
    ]
    .into_iter()
    .zip(1_u32..=12)
    {
        output.set(
            index,
            item.map_or(JsValue::NULL, |item_value| number(f64::from(item_value))),
        );
    }
    Ok(output.into())
}

fn output_numbering_fact(fact: &NumberingHierarchyFact) -> Result<JsValue, String> {
    let output = Array::new_with_length(3);
    output.set(0, usize_number(fact.paragraph_ordinal)?);
    output.set(
        1,
        fact.parent_paragraph_ordinal
            .map(usize_number)
            .transpose()?
            .unwrap_or(JsValue::NULL),
    );
    let children = Array::new();
    for child in &fact.child_paragraph_ordinals {
        children.push(&usize_number(*child)?);
    }
    output.set(2, children.into());
    Ok(output.into())
}

fn output_outline_level_fact(fact: &ParagraphOutlineLevelFact) -> Result<JsValue, String> {
    let output = Array::new_with_length(2);
    output.set(0, usize_number(fact.paragraph_ordinal)?);
    output.set(1, number(f64::from(fact.outline_level)));
    Ok(output.into())
}

fn output_bookmark_fact(fact: &BookmarkFact) -> Result<JsValue, String> {
    let output = Array::new_with_length(4);
    output.set(0, usize_number(fact.paragraph_ordinal)?);
    output.set(1, number(f64::from(fact.bookmark_id)));
    output.set(2, JsValue::from_str(&fact.name));
    output.set(3, output_span(fact.span));
    Ok(output.into())
}

fn output_reference_fact(fact: &InternalReferenceFact) -> Result<JsValue, String> {
    let output = Array::new_with_length(4);
    output.set(0, usize_number(fact.paragraph_ordinal)?);
    output.set(1, JsValue::from_str(&fact.reference_id));
    output.set(
        2,
        JsValue::from_str(match fact.role {
            InternalReferenceRole::Source => "source",
            InternalReferenceRole::Target => "target",
        }),
    );
    output.set(3, output_span(fact.span));
    Ok(output.into())
}

fn output_span(span: StructuralSpan) -> JsValue {
    let output = Array::new_with_length(5);
    output.set(0, number(f64::from(span.start_utf8)));
    output.set(1, number(f64::from(span.end_utf8)));
    output.set(2, number(f64::from(span.start_utf16)));
    output.set(3, number(f64::from(span.end_utf16)));
    output.set(
        4,
        JsValue::from_str(match span.coverage {
            SpanCoverage::Complete => "complete",
            SpanCoverage::ContinuesBefore => "continues-before",
            SpanCoverage::ContinuesAfter => "continues-after",
            SpanCoverage::ContinuesBeforeAndAfter => "continues-before-and-after",
        }),
    );
    output.into()
}

fn number(value: f64) -> JsValue {
    JsValue::from_f64(value)
}

fn usize_number(value: usize) -> Result<JsValue, String> {
    u32::try_from(value)
        .map(|value| number(f64::from(value)))
        .map_err(|_| "projection ordinal exceeds the wire format limit".to_owned())
}

const fn unknown_reason(reason: StructuralFactUnknownReason) -> &'static str {
    match reason {
        StructuralFactUnknownReason::DocumentPartOnly => "document-part-only",
        StructuralFactUnknownReason::StylesPartUnavailable => "styles-part-unavailable",
        StructuralFactUnknownReason::UnsupportedStyles => "unsupported-styles",
        StructuralFactUnknownReason::UnsupportedNumbering => "unsupported-numbering",
        StructuralFactUnknownReason::IncompleteBookmarkRanges => "incomplete-bookmark-ranges",
        StructuralFactUnknownReason::UnsupportedInternalReferences => {
            "unsupported-internal-references"
        }
    }
}

fn output_paragraph_structure(structure: Option<ParagraphStructure>) -> Result<JsValue, String> {
    let output = Array::new();
    if let Some(structure) = structure {
        output.push(&JsValue::from_str("table"));
        output.push(&JsValue::from_str(&format!(
            "table-{}",
            structure.table_ordinal
        )));
        output.push(&usize_number(structure.row)?);
        output.push(&usize_number(structure.column)?);
    }
    Ok(output.into())
}

fn js_error(error: &str) -> JsValue {
    js_sys::Error::new(error).into()
}
