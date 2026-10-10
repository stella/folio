#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::unwrap_used
)]

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::*;
use stella_docx_kernel::{
    DocumentProjection, DocxLimits, FormattingCompleteness, FormattingFactStatus,
    InternalParagraphId, ParagraphIdentityFacts, ProjectionError, project_docx,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("internal-{}", facts.ordinal))
}

struct PackageOptions<'a> {
    namespace: &'a str,
    body: &'a str,
    numbering: &'a str,
}

fn project(
    PackageOptions {
        namespace,
        body,
        numbering,
    }: PackageOptions<'_>,
) -> DocumentProjection {
    let document = format!(
        r#"<w:document xmlns:w="{namespace}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>{body}</w:p></w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{namespace}"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/></w:styles>"#
    );
    let numbering = format!(r#"<w:numbering xmlns:w="{namespace}">{numbering}</w:numbering>"#);
    let relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#;
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, contents) in [
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", styles.as_str()),
        ("word/numbering.xml", numbering.as_str()),
        ("word/_rels/document.xml.rels", relationships),
    ] {
        archive
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        archive.write_all(contents.as_bytes()).unwrap();
    }
    let bytes = archive.finish().unwrap().into_inner();
    project_docx(&bytes, DocxLimits::default(), allocate).unwrap()
}

fn numbering(level_properties: &[String]) -> String {
    let mut xml = String::from(r#"<w:abstractNum w:abstractNumId="1">"#);
    for (level, properties) in level_properties.iter().enumerate() {
        write!(xml, r#"<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>{properties}</w:lvl>"#).unwrap();
    }
    xml.push_str(r#"</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/>"#);
    for (level, properties) in level_properties.iter().enumerate() {
        write!(xml, r#"<w:lvlOverride w:ilvl="{level}"><w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>{properties}</w:lvl></w:lvlOverride>"#).unwrap();
    }
    xml.push_str("</w:num>");
    xml
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

    /// Numbering-label properties do not describe paragraph text, including
    /// properties in level overrides and levels the paragraph does not use.
    #[test]
    fn label_properties_preserve_all_body_facts(
        levels in prop::collection::vec((any::<bool>(), any::<bool>(), any::<bool>()), 1..10),
        runs in prop::collection::vec((prop::sample::select(vec!["a", "é", "😀", "bc"]), any::<bool>()), 1..8),
        strict in any::<bool>(),
    ) {
        let namespace = if strict {
            "http://purl.oclc.org/ooxml/wordprocessingml/main"
        } else {
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        };
        let mut body = String::new();
        for (text, bold) in runs {
            let properties = if bold { "<w:rPr><w:b/></w:rPr>" } else { "" };
            write!(body, "<w:r>{properties}<w:t>{text}</w:t></w:r>").unwrap();
        }
        let properties = levels.iter().map(|(bold, italic, hidden)| {
            let flag = |value| if value { "1" } else { "0" };
            format!(r#"<w:rPr><w:b w:val="{}"/><w:i w:val="{}"/><w:vanish w:val="{}"/><w:rFonts w:ascii="Synthetic"/><w:color w:val="112233"/></w:rPr>"#, flag(*bold), flag(*italic), flag(*hidden))
        }).collect::<Vec<_>>();
        let baseline_numbering = numbering(&vec![String::new(); levels.len()]);
        let styled_numbering = numbering(&properties);
        let baseline = project(PackageOptions { namespace, body: &body, numbering: &baseline_numbering });
        let styled = project(PackageOptions { namespace, body: &body, numbering: &styled_numbering });
        prop_assert_eq!(baseline.formatting_completeness, FormattingCompleteness { bold: FormattingFactStatus::Known, highlight: FormattingFactStatus::Known, superscript: FormattingFactStatus::Known });
        prop_assert_eq!(baseline, styled);
    }
}
