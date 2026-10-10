#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]

use proptest::prelude::{prop_assert_eq, proptest};
use stella_docx_kernel::{
    DocxLimits, FormattingFactStatus, FormattingUnknownReason, InternalParagraphId,
    ParagraphIdentityFacts, ProjectionError, TextFormattingSpan, TextStyle,
    project_main_document_xml,
};

const MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PACKAGE_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/xmlPackage";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const RELATIONSHIP_TYPES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("paragraph-{}", facts.ordinal))
}

fn package(runs: &[(String, bool)], include_styles: bool, table: bool) -> String {
    let mut content = String::new();
    for (text, highlighted) in runs {
        let highlight = if *highlighted { "yellow" } else { "none" };
        content.push_str(&format!(
            r#"<w:r><w:rPr><w:rStyle w:val="Character"/><w:highlight w:val="{highlight}"/></w:rPr><w:t>{text}</w:t></w:r>"#,
        ));
    }
    let paragraph = format!(
        r#"<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:pPr><w:pStyle w:val="Paragraph"/></w:pPr>{content}</w:p>"#,
    );
    let body = if table {
        format!("<w:tbl><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl>")
    } else {
        paragraph
    };
    let document =
        format!(r#"<w:document xmlns:w="{MAIN_NAMESPACE}"><w:body>{body}</w:body></w:document>"#);
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Type="{RELATIONSHIP_TYPES}/officeDocument" Target="parts/content.xml"/></Relationships>"#,
    );
    let part = |name: &str, xml: &str| {
        format!(
            r#"<pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>"#,
        )
    };
    let mut output = format!(r#"<pkg:package xmlns:pkg="{PACKAGE_NAMESPACE}">"#);
    output.push_str(&part("/_rels/.rels", &relationships));
    output.push_str(&part("/parts/content.xml", &document));
    if include_styles {
        let styles = format!(
            r#"<w:styles xmlns:w="{MAIN_NAMESPACE}"><w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Paragraph"><w:rPr><w:highlight w:val="green"/></w:rPr></w:style><w:style w:type="character" w:styleId="Character"><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"#,
        );
        let style_relationships = format!(
            r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Type="{RELATIONSHIP_TYPES}/styles" Target="../definitions/styles.xml"/></Relationships>"#,
        );
        output.push_str(&part("/parts/_rels/content.xml.rels", &style_relationships));
        output.push_str(&part("/definitions/styles.xml", &styles));
    }
    output.push_str("</pkg:package>");
    output
}

// Independent oracle: spans follow authored run lengths and direct highlight flags.
fn expected_highlights(runs: &[(String, bool)]) -> Vec<TextFormattingSpan> {
    let mut spans = Vec::<TextFormattingSpan>::new();
    let mut offset = 0_u32;
    for (text, highlighted) in runs {
        let end = offset
            .checked_add(u32::try_from(text.encode_utf16().count()).expect("bounded fixture text"))
            .expect("bounded fixture offset");
        if *highlighted {
            if let Some(previous) = spans.last_mut().filter(|span| span.end_utf16 == offset) {
                previous.end_utf16 = end;
            } else {
                spans.push(TextFormattingSpan {
                    start_utf16: offset,
                    end_utf16: end,
                    style: TextStyle::Highlight,
                });
            }
        }
        offset = end;
    }
    spans
}

#[test]
fn direct_highlight_in_a_relationship_selected_table_cell_is_known_without_styles() {
    let runs = vec![("Highlighted clause".to_owned(), true)];
    let xml = package(&runs, false, true);
    let projection = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate)
        .expect("bounded table fixture projects");
    assert_eq!(
        projection.formatting_completeness.highlight,
        FormattingFactStatus::Known
    );
    assert_eq!(
        projection.formatting_completeness.bold,
        FormattingFactStatus::Unknown(FormattingUnknownReason::StylesPartUnavailable)
    );
    assert_eq!(
        projection.formatting_completeness.superscript,
        FormattingFactStatus::Unknown(FormattingUnknownReason::StylesPartUnavailable)
    );
    let paragraph = projection
        .paragraphs
        .first()
        .expect("one table-cell paragraph");
    assert_eq!(paragraph.text, "Highlighted clause");
    assert_eq!(
        paragraph
            .package_paragraph_id
            .expect("package paragraph id")
            .value(),
        1
    );
    assert!(paragraph.structure.is_some());
    assert_eq!(paragraph.formatting, expected_highlights(&runs));
}

proptest! {
    #[test]
    fn removing_styles_preserves_direct_highlight_facts_and_completeness(
        runs in proptest::collection::vec(("[a-zA-Zé😀]{1,12}", proptest::bool::ANY), 1..9),
        table in proptest::bool::ANY,
    ) {
        let mut observed = Vec::new();
        for include_styles in [true, false] {
            let xml = package(&runs, include_styles, table);
            let projection = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate)
                .expect("generated bounded package projects");
            prop_assert_eq!(projection.formatting_completeness.highlight, FormattingFactStatus::Known);
            let paragraph = projection.paragraphs.first().expect("generated paragraph");
            let highlights: Vec<_> = paragraph.formatting.iter()
                .filter(|span| span.style == TextStyle::Highlight).cloned().collect();
            prop_assert_eq!(&highlights, &expected_highlights(&runs));
            observed.push(highlights);
            let expected_styles = if include_styles { FormattingFactStatus::Known } else {
                FormattingFactStatus::Unknown(FormattingUnknownReason::StylesPartUnavailable)
            };
            prop_assert_eq!(projection.formatting_completeness.bold, expected_styles);
            prop_assert_eq!(projection.formatting_completeness.superscript, expected_styles);
        }
        prop_assert_eq!(observed.first(), observed.last());
    }
}
