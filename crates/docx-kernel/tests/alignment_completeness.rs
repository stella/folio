#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]

use proptest::prelude::{prop_assert, prop_assert_eq, proptest};
use std::fmt::Write;
use stella_docx_kernel::{
    DocxLimits, FormattingFactStatus, FormattingUnknownReason, InternalParagraphId,
    ParagraphAlignmentFact, ParagraphAlignmentSource, ParagraphAlignmentValue, ParagraphContainer,
    ParagraphIdentityFacts, ProjectionError, project_main_document_xml, project_paragraph_fragment,
};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG: &str = "http://schemas.microsoft.com/office/2006/xmlPackage";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const TYPES: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("paragraph-{}", facts.ordinal))
}

fn package(properties: &[String], styles: Option<&str>, table: bool) -> String {
    let mut body = String::new();
    for property in properties {
        write!(
            body,
            "<w:p><w:pPr>{property}</w:pPr><w:r><w:t>Content</w:t></w:r></w:p>"
        )
        .expect("writing fixture markup to a String cannot fail");
    }
    if table {
        body = format!("<w:tbl><w:tr><w:tc>{body}</w:tc></w:tr></w:tbl>");
    }
    let part = |name: &str, xml: &str| {
        format!(
            r#"<pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>"#
        )
    };
    let root = format!(
        r#"<Relationships xmlns="{RELS}"><Relationship Type="{TYPES}/officeDocument" Target="content/main.xml"/></Relationships>"#
    );
    let document = format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#);
    let mut result = format!(
        r#"<pkg:package xmlns:pkg="{PKG}">{}{}"#,
        part("/_rels/.rels", &root),
        part("/content/main.xml", &document)
    );
    if let Some(styles) = styles {
        let relationships = format!(
            r#"<Relationships xmlns="{RELS}"><Relationship Type="{TYPES}/styles" Target="styles.xml"/></Relationships>"#
        );
        result.push_str(&part("/content/_rels/main.xml.rels", &relationships));
        result.push_str(&part("/content/styles.xml", styles));
    }
    result.push_str("</pkg:package>");
    result
}

fn stylesheet(style_alignment: &str, defaults: &str) -> String {
    format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:pPrDefault><w:pPr>{defaults}</w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Centered"><w:pPr>{style_alignment}</w:pPr></w:style></w:styles>"#
    )
}

fn alignment(index: u8) -> (&'static str, ParagraphAlignmentValue) {
    match index {
        0 => ("left", ParagraphAlignmentValue::Left),
        1 => ("center", ParagraphAlignmentValue::Center),
        2 => ("right", ParagraphAlignmentValue::Right),
        3 => ("both", ParagraphAlignmentValue::Justify),
        _ => panic!("fixture alignment index outside declared values"),
    }
}

proptest! {
    #[test]
    fn removing_styles_preserves_direct_alignment_and_never_claims_inherited_absence(
        inherited in 0_u8..4,
        direct in proptest::collection::vec(proptest::option::of(0_u8..4), 1..8),
        table in proptest::bool::ANY,
    ) {
        let (inherited_token, inherited_value) = alignment(inherited);
        let properties = direct.iter().map(|value| {
            let override_markup = value.map_or_else(String::new, |index| {
                let (token, _) = alignment(index);
                format!(r#"<w:jc w:val="{token}"/>"#)
            });
            format!(r#"<w:pStyle w:val="Centered"/>{override_markup}"#)
        }).collect::<Vec<_>>();
        let styles = stylesheet(&format!(r#"<w:jc w:val="{inherited_token}"/>"#), "");
        let full = project_main_document_xml(package(&properties, Some(&styles), table).as_bytes(), DocxLimits::default(), allocate).unwrap();
        let missing = project_main_document_xml(package(&properties, None, table).as_bytes(), DocxLimits::default(), allocate).unwrap();
        prop_assert_eq!(full.formatting_completeness.alignment, FormattingFactStatus::Known);
        let expected_status = if direct.iter().all(Option::is_some) { FormattingFactStatus::Known }
            else { FormattingFactStatus::Unknown(FormattingUnknownReason::StylesPartUnavailable) };
        prop_assert_eq!(missing.formatting_completeness.alignment, expected_status);
        let expected_container = if table { ParagraphContainer::TableCell } else { ParagraphContainer::Body };
        if direct.len() == 1 {
            let fragment = project_paragraph_fragment(package(&properties, None, table).as_bytes(), DocxLimits::default(), allocate).unwrap();
            prop_assert_eq!(fragment.formatting_completeness.alignment, expected_status);
            prop_assert_eq!(fragment.paragraphs.first().unwrap().container, expected_container);
            prop_assert!(fragment.paragraphs.first().unwrap().structure.is_none());
        }
        for ((full_paragraph, missing_paragraph), authored) in full.paragraphs.iter().zip(&missing.paragraphs).zip(direct) {
            let expected = authored.map_or(
                ParagraphAlignmentFact { value: inherited_value, source: ParagraphAlignmentSource::Style },
                |index| ParagraphAlignmentFact { value: alignment(index).1, source: ParagraphAlignmentSource::Direct },
            );
            prop_assert_eq!(full_paragraph.container, expected_container);
            prop_assert_eq!(missing_paragraph.container, expected_container);
            prop_assert_eq!(full_paragraph.alignment, Some(expected));
            prop_assert_eq!(missing_paragraph.alignment, authored.map(|index| ParagraphAlignmentFact {
                value: alignment(index).1, source: ParagraphAlignmentSource::Direct,
            }));
        }
    }
}

#[test]
fn fragment_alignment_separates_defaults_absence_and_unread_values() {
    let cases = [
        ("", "", "", None, FormattingFactStatus::Known),
        (
            "",
            "",
            r#"<w:jc w:val="center"/>"#,
            Some(ParagraphAlignmentFact {
                value: ParagraphAlignmentValue::Center,
                source: ParagraphAlignmentSource::DocDefaults,
            }),
            FormattingFactStatus::Known,
        ),
        (
            r#"<w:jc w:val="unsupported"/>"#,
            "",
            "",
            None,
            FormattingFactStatus::Unknown(FormattingUnknownReason::UnsupportedAlignment),
        ),
    ];
    for table in [false, true] {
        for (direct, style, defaults, expected, status) in cases {
            let styles = stylesheet(style, defaults);
            let xml = package(&[direct.to_owned()], Some(&styles), table);
            let fragment =
                project_paragraph_fragment(xml.as_bytes(), DocxLimits::default(), allocate)
                    .unwrap();
            assert_eq!(fragment.formatting_completeness.alignment, status);
            let [paragraph] = fragment.paragraphs.as_slice() else {
                panic!("fragment fixture must emit exactly one paragraph")
            };
            assert_eq!(paragraph.alignment, expected);
            assert!(paragraph.structure.is_none());
            assert_eq!(
                paragraph.container,
                if table {
                    ParagraphContainer::TableCell
                } else {
                    ParagraphContainer::Body
                }
            );
        }
    }
}

proptest! {
    #[test]
    fn paragraph_alignment_does_not_require_a_shadowed_table_style(
        selected in 0_u8..4,
        cyclic in proptest::bool::ANY,
    ) {
        let (token, value) = alignment(selected);
        let table_definition = if cyclic {
            r#"<w:basedOn w:val="Unread"/>"#
        } else {
            r#"<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="right"/></w:pPr></w:tblStylePr>"#
        };
        let styles = format!(r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Centered"><w:pPr><w:jc w:val="{token}"/></w:pPr></w:style><w:style w:type="table" w:default="1" w:styleId="Unread">{table_definition}</w:style></w:styles>"#);
        let xml = package(&[r#"<w:pStyle w:val="Centered"/>"#.to_owned()], Some(&styles), true);
        let projection = project_paragraph_fragment(xml.as_bytes(), DocxLimits::default(), allocate).unwrap();
        prop_assert_eq!(projection.formatting_completeness.alignment, FormattingFactStatus::Known);
        prop_assert_eq!(projection.paragraphs.first().unwrap().alignment, Some(ParagraphAlignmentFact {
            value, source: ParagraphAlignmentSource::Style,
        }));
    }
}
