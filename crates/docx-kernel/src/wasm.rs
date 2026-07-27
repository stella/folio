use crate::{
    BookmarkFact, DocumentProjection, DocumentStructureFacts, DocxLimits, InternalParagraphId,
    InternalReferenceFact, InternalReferenceRole, NumberingHierarchyFact, ParagraphIdentityFacts,
    ParagraphIndentationFact, ParagraphStructure, ProjectedParagraph, RevisionProjectionStatus,
    RevisionUnsupportedReason, SpanCoverage, StructuralFactSet, StructuralFactUnknownReason,
    StructuralSpan, TextStyle, project_docx,
};
use js_sys::Array;
use wasm_bindgen::{JsCast, prelude::*};

const DOCX_PROJECTION_SCHEMA_VERSION: u32 = 2;

#[wasm_bindgen(typescript_custom_section)]
const TYPESCRIPT_TYPES: &str = r#"
export type DocxProjectionFormattingSpan = readonly [
  startUtf16: number,
  endUtf16: number,
  style: "bold" | "highlight",
];
export type DocxProjectionStructure =
  | readonly []
  | readonly [
      type: "table",
      tableId: string,
      row: number,
      column: number,
    ];
export type DocxProjectionParagraph = readonly [
  ordinal: number,
  text: string,
  packageParagraphId: string | null,
  formatting: readonly DocxProjectionFormattingSpan[],
  structure: DocxProjectionStructure,
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
  schemaVersion: 2,
  paragraphs: readonly DocxProjectionParagraph[],
  structuralFacts: DocxProjectionStructuralFacts,
  revisionStatus: DocxProjectionRevisionStatus,
];
"#;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(typescript_type = "DocxProjectionWire")]
    pub type DocxProjectionWire;
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

fn project_docx_projection(bytes: &[u8]) -> Result<DocumentProjection, String> {
    let limits = DocxLimits::default();
    project_docx(bytes, limits, |facts: ParagraphIdentityFacts<'_>| {
        InternalParagraphId::new(format!("projected-{}", facts.ordinal))
    })
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
    let output = Array::new_with_length(5);
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
        output_span.set(
            2,
            JsValue::from_str(match span.style {
                TextStyle::Bold => "bold",
                TextStyle::Highlight => "highlight",
            }),
        );
        formatting.push(&output_span);
    }
    output.set(3, formatting.into());
    output.set(4, output_paragraph_structure(paragraph.structure.clone())?);
    Ok(output.into())
}

fn output_projection_with_structure(projection: &DocumentProjection) -> Result<JsValue, String> {
    let output = Array::new_with_length(4);
    output.set(
        0,
        JsValue::from_f64(f64::from(DOCX_PROJECTION_SCHEMA_VERSION)),
    );
    output.set(1, output_paragraphs(projection)?);
    output.set(2, output_structural_facts(&projection.structural_facts)?);
    output.set(3, output_revision_status(&projection.revision_status));
    Ok(output.into())
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
    let output = Array::new_with_length(4);
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
