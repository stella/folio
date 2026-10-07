#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used
)]

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::{prop_assert, prop_assert_eq, proptest};
use stella_docx_kernel::{
    DocxLimits, FormattingProjectionStatus, FormattingUnknownReason, InternalParagraphId,
    StructuralFactSet, StructuralFactUnknownReason, TextFormattingSpan, TextStyle, project_docx,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn package(document: &str, styles: &str) -> Vec<u8> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    let relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    for (name, xml) in [
        ("word/document.xml", document),
        ("word/styles.xml", styles),
        ("word/_rels/document.xml.rels", relationships),
    ] {
        archive
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        archive.write_all(xml.as_bytes()).unwrap();
    }
    archive.finish().unwrap().into_inner()
}

fn styles_xml(definitions: &str, default_bold: bool) -> String {
    format!(
        r#"<w:styles xmlns:w="{NAMESPACE}"><w:docDefaults><w:rPrDefault><w:rPr><w:b w:val="{default_bold}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:ind w:start="100"/><w:outlineLvl w:val="2"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Default"><w:rPr><w:b/></w:rPr><w:pPr><w:ind w:start="300"/><w:outlineLvl w:val="4"/></w:pPr></w:style>{definitions}</w:styles>"#
    )
}

fn document_xml(paragraph_style: &str, character_style: &str) -> String {
    format!(
        r#"<w:document xmlns:w="{NAMESPACE}"><w:body><w:p><w:pPr>{paragraph_style}</w:pPr><w:r><w:rPr>{character_style}</w:rPr><w:t>a😀é</w:t></w:r></w:p></w:body></w:document>"#
    )
}

fn chain(kind: &str, prefix: &str, toggles: &[bool], cyclic: bool) -> String {
    let mut definitions = String::new();
    for (index, toggle) in toggles.iter().enumerate() {
        let base = if index == 0 {
            if cyclic {
                format!("{prefix}{}", toggles.len() - 1)
            } else {
                "MissingBase".to_owned()
            }
        } else {
            format!("{prefix}{}", index - 1)
        };
        let properties = if kind == "paragraph" {
            format!(
                r#"<w:pPr><w:ind w:start="{}"/><w:outlineLvl w:val="1"/></w:pPr>"#,
                500 + index
            )
        } else {
            String::new()
        };
        write!(definitions, r#"<w:style w:type="{kind}" w:styleId="{prefix}{index}"><w:basedOn w:val="{base}"/><w:rPr><w:b w:val="{toggle}"/></w:rPr>{properties}</w:style>"#).unwrap();
    }
    definitions
}

proptest! {
    #[test]
    fn style_chains_and_missing_selections_have_effective_facts(
        paragraph_toggles in proptest::collection::vec(proptest::bool::ANY, 1..17),
        character_toggles in proptest::collection::vec(proptest::bool::ANY, 1..17),
        default_bold in proptest::bool::ANY,
        paragraph_selection in 0u8..3,
        character_missing in proptest::bool::ANY,
    ) {
        let definitions = format!("{}{}", chain("paragraph", "P", &paragraph_toggles, false), chain("character", "C", &character_toggles, false));
        let paragraph_style = match paragraph_selection {
            0 => String::new(),
            1 => r#"<w:pStyle w:val="MissingParagraph"/>"#.to_owned(),
            _ => format!(r#"<w:pStyle w:val="P{}"/>"#, paragraph_toggles.len() - 1),
        };
        let character_style = if character_missing {
            r#"<w:rStyle w:val="MissingCharacter"/>"#.to_owned()
        } else {
            format!(r#"<w:rStyle w:val="C{}"/>"#, character_toggles.len() - 1)
        };
        let projection = project_docx(
            &package(&document_xml(&paragraph_style, &character_style), &styles_xml(&definitions, default_bold)),
            DocxLimits::default(),
            |facts| InternalParagraphId::new(format!("paragraph-{}", facts.ordinal)),
        ).unwrap();
        prop_assert_eq!(projection.formatting_status, FormattingProjectionStatus::Complete);
        let paragraph_bold = match paragraph_selection {
            0 => true,
            1 => false,
            _ => paragraph_toggles.iter().filter(|value| **value).count() % 2 == 1,
        };
        let character_bold = !character_missing && character_toggles.iter().filter(|value| **value).count() % 2 == 1;
        let expected_bold = default_bold ^ paragraph_bold ^ character_bold;
        let expected_formatting = if expected_bold {
            vec![TextFormattingSpan { start_utf16: 0, end_utf16: 4, style: TextStyle::Bold }]
        } else {
            Vec::new()
        };
        prop_assert_eq!(&projection.paragraphs.first().unwrap().formatting, &expected_formatting);
        let StructuralFactSet::Known(indentation) = &projection.structural_facts.indentation else {
            prop_assert!(false, "bounded style chains keep indentation known");
            return Ok(());
        };
        let expected_start = match paragraph_selection {
            0 => 300,
            1 => 100,
            _ => i32::try_from(499 + paragraph_toggles.len()).unwrap(),
        };
        prop_assert_eq!(indentation.first().unwrap().value.start_twips, Some(expected_start));
        let StructuralFactSet::Known(outline) = &projection.structural_facts.outline_levels else {
            prop_assert!(false, "bounded style chains keep outline levels known");
            return Ok(());
        };
        prop_assert_eq!(outline.first().unwrap().outline_level, match paragraph_selection { 0 => 4, 1 => 2, _ => 1 });
        prop_assert_eq!(projection.structural_facts.numbering_hierarchy, StructuralFactSet::Known(Vec::new()));
    }

    #[test]
    fn selected_cycles_remain_typed_unknown_and_unused_cycles_do_not_interfere(
        toggles in proptest::collection::vec(proptest::bool::ANY, 2..17),
        character_cycle in proptest::bool::ANY,
        selected in proptest::bool::ANY,
    ) {
        let (kind, prefix) = if character_cycle { ("character", "C") } else { ("paragraph", "P") };
        let style = if selected { format!("{prefix}{}", toggles.len() - 1) } else { "Missing".to_owned() };
        let paragraph_style = if character_cycle { String::new() } else { format!(r#"<w:pStyle w:val="{style}"/>"#) };
        let character_style = if character_cycle { format!(r#"<w:rStyle w:val="{style}"/>"#) } else { String::new() };
        let projection = project_docx(
            &package(&document_xml(&paragraph_style, &character_style), &styles_xml(&chain(kind, prefix, &toggles, true), false)),
            DocxLimits::default(),
            |facts| InternalParagraphId::new(format!("paragraph-{}", facts.ordinal)),
        ).unwrap();
        prop_assert_eq!(projection.formatting_status, if selected {
            FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
        } else { FormattingProjectionStatus::Complete });
        if selected && !character_cycle {
            prop_assert_eq!(projection.structural_facts.indentation, StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
            prop_assert_eq!(projection.structural_facts.outline_levels, StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
            prop_assert_eq!(projection.structural_facts.numbering_hierarchy, StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
        } else {
            prop_assert!(matches!(projection.structural_facts.indentation, StructuralFactSet::Known(_)));
            prop_assert!(matches!(projection.structural_facts.outline_levels, StructuralFactSet::Known(_)));
            prop_assert_eq!(projection.structural_facts.numbering_hierarchy, StructuralFactSet::Known(Vec::new()));
        }
    }
}
