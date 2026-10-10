#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::*;
use stella_docx_kernel::{
    DocumentProjection, DocxLimits, FormattingCompleteness, FormattingFactStatus,
    FormattingUnknownReason, InternalParagraphId, ParagraphAlignmentFact, ParagraphAlignmentSource,
    ParagraphAlignmentValue, StructuralFactSet, StructuralFactUnknownReason, TextFormattingSpan,
    TextStyle, project_docx,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TableSelection {
    Explicit,
    Default,
    Missing,
}

const NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn project(body: &str, definitions: &str) -> DocumentProjection {
    let document =
        format!(r#"<w:document xmlns:w="{NAMESPACE}"><w:body>{body}</w:body></w:document>"#);
    let styles = format!(r#"<w:styles xmlns:w="{NAMESPACE}">{definitions}</w:styles>"#);
    let relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, contents) in [
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", styles.as_str()),
        ("word/_rels/document.xml.rels", relationships),
    ] {
        archive
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        archive.write_all(contents.as_bytes()).unwrap();
    }
    project_docx(
        &archive.finish().unwrap().into_inner(),
        DocxLimits::default(),
        |facts| InternalParagraphId::new(format!("paragraph-{}", facts.ordinal)),
    )
    .unwrap()
}

fn table(style: &str, content: &str) -> String {
    let properties = if style.is_empty() {
        String::new()
    } else {
        format!(r#"<w:tblStyle w:val="{style}"/>"#)
    };
    format!("<w:tbl><w:tblPr>{properties}</w:tblPr><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>")
}

fn chain(toggles: &[bool], cyclic: bool) -> String {
    let mut styles = String::new();
    for (index, toggle) in toggles.iter().enumerate() {
        let parent = if index == 0 {
            if cyclic {
                format!("T{}", toggles.len() - 1)
            } else {
                "MissingBase".to_owned()
            }
        } else {
            format!("T{}", index - 1)
        };
        write!(styles, r#"<w:style w:type="table" w:styleId="T{index}"><w:basedOn w:val="{parent}"/><w:rPr><w:b w:val="{toggle}"/></w:rPr><w:pPr><w:ind w:start="{}"/><w:outlineLvl w:val="1"/><w:jc w:val="center"/></w:pPr></w:style>"#, 500 + index).unwrap();
    }
    styles
}

fn default_properties(bold: bool) -> String {
    format!(
        r#"<w:docDefaults><w:rPrDefault><w:rPr><w:b w:val="{bold}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:ind w:start="100"/><w:outlineLvl w:val="4"/><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>"#
    )
}

fn bold_spans(text: &str, bold: bool) -> Vec<TextFormattingSpan> {
    if !bold {
        return Vec::new();
    }
    vec![TextFormattingSpan {
        start_utf16: 0,
        end_utf16: u32::try_from(text.encode_utf16().count()).unwrap(),
        style: TextStyle::Bold,
    }]
}

const fn alignment(
    value: ParagraphAlignmentValue,
    source: ParagraphAlignmentSource,
) -> ParagraphAlignmentFact {
    ParagraphAlignmentFact { value, source }
}

fn config() -> ProptestConfig {
    let factor = std::env::var("PROPERTY_TEST_NUM_RUNS_FACTOR")
        .ok()
        .and_then(|value| value.parse::<u32>().ok())
        .filter(|factor| *factor > 0)
        .unwrap_or(1);
    ProptestConfig {
        cases: 64 * factor,
        ..ProptestConfig::default()
    }
}

proptest! {
    #![proptest_config(config())]

    /// A table chain contributes its own cascade tier; paragraph and
    /// character style toggles compose, while direct run values override.
    #[test]
    fn table_chain_properties_follow_cascade_tiers(
        toggles in prop::collection::vec(any::<bool>(), 1..17),
        default_bold in any::<bool>(),
        paragraph_toggle in proptest::option::of(any::<bool>()),
        character_toggle in proptest::option::of(any::<bool>()),
        direct_bold in proptest::option::of(any::<bool>()),
        direct_paragraph in any::<bool>(),
        selection in prop::sample::select(vec![TableSelection::Explicit, TableSelection::Default, TableSelection::Missing]),
        text in prop::sample::select(vec!["a😀é", "😀😀", "ébc", "x"]),
    ) {
        let table_definitions = chain(&toggles, false).replace(
            &format!(r#"w:styleId="T{}""#, toggles.len() - 1),
            &format!(r#"w:styleId="T{}" w:default="1""#, toggles.len() - 1),
        );
        let mut definitions = format!("{}{}", default_properties(default_bold), table_definitions);
        let paragraph_style = paragraph_toggle.map_or_else(String::new, |toggle| {
            write!(definitions, r#"<w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b w:val="{toggle}"/></w:rPr><w:pPr><w:ind w:start="700"/><w:outlineLvl w:val="2"/><w:jc w:val="both"/></w:pPr></w:style>"#).unwrap();
            r#"<w:pStyle w:val="P"/>"#.to_owned()
        });
        let character_style = character_toggle.map_or_else(String::new, |toggle| {
            write!(definitions, r#"<w:style w:type="character" w:styleId="C"><w:rPr><w:b w:val="{toggle}"/></w:rPr></w:style>"#).unwrap();
            r#"<w:rStyle w:val="C"/>"#.to_owned()
        });
        let direct = direct_bold.map_or_else(String::new, |bold| format!(r#"<w:b w:val="{bold}"/>"#));
        let direct_properties = if direct_paragraph { r#"<w:ind w:start="900"/><w:outlineLvl w:val="3"/><w:jc w:val="right"/>"# } else { "" };
        let paragraph = format!("<w:p><w:pPr>{paragraph_style}{direct_properties}</w:pPr><w:r><w:rPr>{character_style}{direct}</w:rPr><w:t>{text}</w:t></w:r></w:p>");
        let selected = match selection {
            TableSelection::Explicit => format!("T{}", toggles.len() - 1),
            TableSelection::Default => String::new(),
            TableSelection::Missing => "MissingTable".to_owned(),
        };
        let missing_table = selection == TableSelection::Missing;
        let projection = project(&table(&selected, &paragraph), &definitions);
        prop_assert_eq!(projection.formatting_completeness, FormattingCompleteness { bold: FormattingFactStatus::Known, highlight: FormattingFactStatus::Known, superscript: FormattingFactStatus::Known });
        let table_toggle = !missing_table && toggles.iter().filter(|value| **value).count() % 2 == 1;
        let expected_bold = direct_bold.unwrap_or_else(|| default_bold ^ table_toggle ^ paragraph_toggle.unwrap_or(false) ^ character_toggle.unwrap_or(false));
        prop_assert_eq!(&projection.paragraphs[0].formatting, &bold_spans(text, expected_bold));
        prop_assert_eq!(projection.paragraphs[0].text.as_str(), text);
        let (indent, outline, value, source) = if direct_paragraph {
            (900, 3, ParagraphAlignmentValue::Right, ParagraphAlignmentSource::Direct)
        } else if paragraph_toggle.is_some() {
            (700, 2, ParagraphAlignmentValue::Justify, ParagraphAlignmentSource::Style)
        } else if missing_table {
            (100, 4, ParagraphAlignmentValue::Left, ParagraphAlignmentSource::Style)
        } else {
            (i32::try_from(499 + toggles.len()).unwrap(), 1, ParagraphAlignmentValue::Center, ParagraphAlignmentSource::Style)
        };
        let StructuralFactSet::Known(indentation) = &projection.structural_facts.indentation else {
            prop_assert!(false, "bounded table chains keep indentation Known");
            return Ok(());
        };
        let StructuralFactSet::Known(outlines) = &projection.structural_facts.outline_levels else {
            prop_assert!(false, "bounded table chains keep outline Known");
            return Ok(());
        };
        prop_assert_eq!(indentation[0].value.start_twips, Some(indent));
        prop_assert_eq!(outlines[0].outline_level, outline);
        prop_assert_eq!(projection.paragraphs[0].alignment, Some(alignment(value, source)));
    }

    /// A table cycle affects only projections selecting that table style;
    /// selecting a missing definition adds no table tier.
    #[test]
    fn selected_cycles_are_typed_unknown_and_unused_cycles_preserve_facts(
        toggles in prop::collection::vec(any::<bool>(), 2..17),
        selection in prop::sample::select(vec![TableSelection::Explicit, TableSelection::Default, TableSelection::Missing]),
    ) {
        let name = match selection {
            TableSelection::Explicit => format!("T{}", toggles.len() - 1),
            TableSelection::Default => String::new(),
            TableSelection::Missing => "MissingTable".to_owned(),
        };
        let paragraph = "<w:p><w:r><w:t>é😀</w:t></w:r></w:p>";
        let table_definitions = chain(&toggles, true).replace(
            &format!(r#"w:styleId="T{}""#, toggles.len() - 1),
            &format!(r#"w:styleId="T{}" w:default="1""#, toggles.len() - 1),
        );
        let definitions = format!("{}{}", default_properties(false), table_definitions);
        let projection = project(&table(&name, paragraph), &definitions);
        if selection == TableSelection::Missing {
            let plain = project(&table("", paragraph), &default_properties(false));
            prop_assert_eq!(projection, plain);
        } else {
            prop_assert_eq!(projection.formatting_completeness, FormattingCompleteness { bold: FormattingFactStatus::Unknown(FormattingUnknownReason::UnsupportedStyles), highlight: FormattingFactStatus::Known, superscript: FormattingFactStatus::Unknown(FormattingUnknownReason::UnsupportedStyles) });
            prop_assert_eq!(&projection.structural_facts.indentation, &StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
            prop_assert_eq!(&projection.structural_facts.outline_levels, &StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
            prop_assert_eq!(&projection.structural_facts.numbering_hierarchy, &StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles));
        }
    }

    /// Nested tables choose their nearest table's style; an omitted inner
    /// style selects the default without inheriting outer table properties.
    #[test]
    fn nested_tables_isolate_their_style_selection(
        outer_bold in any::<bool>(),
        inner_bold in proptest::option::of(any::<bool>()),
        default_table_bold in any::<bool>(),
    ) {
        let mut definitions = format!(r#"{}<w:style w:type="table" w:styleId="Outer"><w:rPr><w:b w:val="{outer_bold}"/></w:rPr><w:pPr><w:ind w:start="500"/></w:pPr></w:style>"#, default_properties(false));
        write!(definitions, r#"<w:style w:type="table" w:styleId="DefaultTable" w:default="1"><w:rPr><w:b w:val="{default_table_bold}"/></w:rPr><w:pPr><w:ind w:start="600"/></w:pPr></w:style>"#).unwrap();
        let inner_style = inner_bold.map_or("", |bold| {
            write!(definitions, r#"<w:style w:type="table" w:styleId="Inner"><w:rPr><w:b w:val="{bold}"/></w:rPr><w:pPr><w:ind w:start="700"/></w:pPr></w:style>"#).unwrap();
            "Inner"
        });
        let paragraph = "<w:p><w:r><w:t>é😀</w:t></w:r></w:p>";
        let inner = table(inner_style, paragraph);
        let body = format!("{}{}", table("Outer", &format!("{paragraph}{inner}{paragraph}")), paragraph);
        let projection = project(&body, &definitions);
        prop_assert_eq!(projection.formatting_completeness, FormattingCompleteness { bold: FormattingFactStatus::Known, highlight: FormattingFactStatus::Known, superscript: FormattingFactStatus::Known });
        prop_assert_eq!(projection.paragraphs.len(), 4);
        for index in [0, 2] {
            prop_assert_eq!(&projection.paragraphs[index].formatting, &bold_spans("é😀", outer_bold));
        }
        prop_assert_eq!(&projection.paragraphs[1].formatting, &bold_spans("é😀", inner_bold.unwrap_or(default_table_bold)));
        prop_assert_eq!(&projection.paragraphs[3].formatting, &bold_spans("é😀", false));
        let StructuralFactSet::Known(indentation) = &projection.structural_facts.indentation else {
            prop_assert!(false, "nested table indentation remains Known");
            return Ok(());
        };
        prop_assert_eq!(indentation.iter().map(|fact| fact.value.start_twips).collect::<Vec<_>>(), vec![Some(500), Some(if inner_bold.is_some() { 700 } else { 600 }), Some(500), Some(100)]);
    }
}

#[test]
fn conditional_projected_properties_are_explicitly_unsupported() {
    let paragraph = "<w:p><w:r><w:t>é😀</w:t></w:r></w:p>";
    for properties in [
        "<w:rPr><w:b/></w:rPr>",
        r#"<w:pPr><w:ind w:start="500"/></w:pPr>"#,
        r#"<w:pPr><w:outlineLvl w:val="1"/></w:pPr>"#,
        r#"<w:pPr><w:jc w:val="center"/></w:pPr>"#,
    ] {
        let definitions = format!(
            r#"{}<w:style w:type="table" w:styleId="Conditional"><w:tblStylePr w:type="firstRow">{properties}</w:tblStylePr></w:style>"#,
            default_properties(false)
        );
        let projection = project(&table("Conditional", paragraph), &definitions);
        assert_eq!(
            projection.formatting_completeness,
            FormattingCompleteness {
                bold: FormattingFactStatus::Unknown(FormattingUnknownReason::UnsupportedStyles),
                highlight: FormattingFactStatus::Known,
                superscript: FormattingFactStatus::Unknown(
                    FormattingUnknownReason::UnsupportedStyles
                )
            }
        );
        assert_eq!(
            projection.structural_facts.indentation,
            StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
        );
    }
}

#[test]
fn conditional_fonts_do_not_affect_projected_properties() {
    let paragraph = "<w:p><w:r><w:t>é😀</w:t></w:r></w:p>";
    let base = format!(
        r#"{}<w:style w:type="table" w:styleId="Conditional"/>"#,
        default_properties(false)
    );
    let conditional = format!(
        r#"{}<w:style w:type="table" w:styleId="Conditional"><w:tblStylePr w:type="firstRow"><w:rPr><w:rFonts w:ascii="Synthetic" w:hAnsi="Synthetic"/></w:rPr></w:tblStylePr></w:style>"#,
        default_properties(false)
    );
    assert_eq!(
        project(&table("Conditional", paragraph), &base),
        project(&table("Conditional", paragraph), &conditional)
    );
}
