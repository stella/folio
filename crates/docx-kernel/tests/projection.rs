#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]
// The integration fixtures intentionally use assertion-style failure and index
// only after constructing or checking the exact collection shape under test.

use std::collections::HashSet;
use std::io::{Cursor, Write};

use stella_docx_kernel::{
    DocxLimits, InternalParagraphId, InternalReferenceRole, PackageParagraphId,
    ParagraphIdentityFacts, ParagraphStructure, ProjectionError, ProjectionOptions,
    RevisionProjectionStatus, RevisionUnsupportedReason, RevisionView, SpanCoverage,
    StructuralFactSet, StructuralFactUnknownReason, TextFormattingSpan, TextMaterialization,
    TextStyle, extract_document_parts, extract_document_xml, project_document_xml,
    project_document_xml_with_options, project_docx, project_docx_with_options,
};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const BASIC_XML: &[u8] = include_bytes!("../fixtures/basic.xml");
const COMMENT_XML: &[u8] = include_bytes!("../fixtures/comment.xml");
const HIDDEN_XML: &[u8] = include_bytes!("../fixtures/hidden.xml");
const MISSING_ID_XML: &[u8] = include_bytes!("../fixtures/missingParaId.xml");
const PLACEHOLDER_XML: &[u8] = include_bytes!("../fixtures/placeholder.xml");
const TEXTBOX_XML: &[u8] = include_bytes!("../fixtures/textbox.xml");
const SPECIAL_VISIBLE_TEXT: &str = "\t\t\u{000c}\u{000e}\u{000b}\r\u{001f}\u{001e}©";

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("internal-{}", facts.ordinal))
}

fn texts(xml: &[u8]) -> Vec<String> {
    project_document_xml(xml, allocate)
        .expect("fixture projection should succeed")
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

fn texts_in_view(xml: &[u8], revision_view: RevisionView) -> Vec<String> {
    project_document_xml_with_options(
        xml,
        ProjectionOptions {
            revision_view,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .expect("fixture projection should succeed")
    .paragraphs
    .into_iter()
    .map(|paragraph| paragraph.text)
    .collect()
}

#[test]
fn projects_current_and_original_inline_revision_views() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:r><w:t>A</w:t></w:r>
      <w:del><w:r><w:delText>old</w:delText></w:r></w:del>
      <w:ins><w:r><w:t>new</w:t></w:r></w:ins>
      <w:moveFrom><w:r><w:t>source</w:t></w:r></w:moveFrom>
      <w:moveTo><w:r><w:t>destination</w:t></w:r></w:moveTo>
      <w:r><w:t>Z</w:t></w:r>
    </w:p></w:body></w:document>"#;

    assert_eq!(
        texts_in_view(xml, RevisionView::Current),
        ["AnewdestinationZ"]
    );
    assert_eq!(texts_in_view(xml, RevisionView::Original), ["AoldsourceZ"]);
    for view in [RevisionView::Current, RevisionView::Original] {
        let projection = project_document_xml_with_options(
            xml,
            ProjectionOptions {
                revision_view: view,
                ..ProjectionOptions::default()
            },
            allocate,
        )
        .unwrap();
        assert_eq!(
            projection.revision_status,
            RevisionProjectionStatus::Complete
        );
    }
}

#[test]
fn resolves_paragraph_mark_revisions_without_losing_offsets_or_identity() {
    let deleted_break = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
      <w:p w14:paraId="00000001"><w:pPr><w:rPr><w:del w:id="7"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>
      </w:p>
      <w:p w14:paraId="00000002"><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>😀B</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let current = project_document_xml_with_options(
        deleted_break.as_bytes(),
        ProjectionOptions {
            revision_view: RevisionView::Current,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .unwrap();

    assert_eq!(current.revision_status, RevisionProjectionStatus::Complete);
    assert_eq!(current.paragraphs.len(), 1);
    assert_eq!(current.paragraphs[0].text, "A😀B");
    assert_eq!(
        current.paragraphs[0].package_paragraph_id,
        PackageParagraphId::parse("00000001")
    );
    assert_eq!(
        current.paragraphs[0].formatting,
        [
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 1,
                end_utf16: 4,
                style: TextStyle::Highlight,
            },
        ]
    );
    assert_eq!(
        texts_in_view(deleted_break.as_bytes(), RevisionView::Original),
        ["A", "😀B"]
    );

    let inserted_break = deleted_break.replace("w:del", "w:ins");
    assert_eq!(
        texts_in_view(inserted_break.as_bytes(), RevisionView::Original),
        ["A😀B"]
    );
}

#[test]
fn exposes_incomplete_revision_capabilities_instead_of_guessing() {
    let property_change = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:pPr><w:ind w:left="720"/><w:pPrChange w:id="1"><w:pPr><w:ind w:left="1440"/></w:pPr></w:pPrChange></w:pPr>
      <w:r><w:t>Text</w:t></w:r>
    </w:p></w:body></w:document>"#;
    let package = package_with_minimal_styles(property_change);
    let current = project_docx_with_options(
        &package,
        DocxLimits::default(),
        ProjectionOptions {
            revision_view: RevisionView::Current,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .unwrap();
    assert_eq!(current.revision_status, RevisionProjectionStatus::Complete);
    let StructuralFactSet::Known(indentation) = current.structural_facts.indentation else {
        panic!("current indentation should be known");
    };
    assert_eq!(indentation[0].value.left_twips, Some(720));

    let original = project_docx_with_options(
        &package,
        DocxLimits::default(),
        ProjectionOptions {
            revision_view: RevisionView::Original,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .unwrap();
    assert_eq!(
        original.revision_status,
        RevisionProjectionStatus::Incomplete(vec![
            RevisionUnsupportedReason::UnsupportedRevisionMarkup
        ])
    );

    let structural_table_revision = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr>
      <w:trPr><w:del w:id="2"/></w:trPr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
    </w:tr></w:tbl></w:body></w:document>"#;
    let projection = project_document_xml(structural_table_revision, allocate).unwrap();
    assert_eq!(
        projection.revision_status,
        RevisionProjectionStatus::Incomplete(vec![
            RevisionUnsupportedReason::StructuralTableRevision
        ])
    );
}

#[test]
fn matches_visible_text_semantics_of_existing_ooxml_fixtures() {
    assert_eq!(
        texts(BASIC_XML),
        vec![
            format!("test{SPECIAL_VISIBLE_TEXT}"),
            String::new(),
            String::new(),
            String::new(),
        ]
    );
    assert_eq!(texts(COMMENT_XML), vec!["test", ""]);
    assert_eq!(texts(MISSING_ID_XML), vec!["", ""]);
    assert_eq!(texts(TEXTBOX_XML), vec!["test", ""]);
    assert_eq!(
        texts(PLACEHOLDER_XML),
        vec![
            "testtest".to_owned(),
            format!("testnot placeholder{SPECIAL_VISIBLE_TEXT}test"),
            String::new(),
        ]
    );
    assert_eq!(
        texts(HIDDEN_XML),
        vec![
            String::new(),
            String::new(),
            String::new(),
            format!("test{SPECIAL_VISIBLE_TEXT}"),
            String::new(),
        ]
    );
}

#[test]
fn respects_explicit_false_on_off_values() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>visible zero</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:vanish w:val="false"/></w:rPr><w:t>visible false</w:t></w:r></w:p>
      <w:sdt><w:sdtPr><w:showingPlcHdr w:val="off"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>real content</w:t></w:r></w:p></w:sdtContent></w:sdt>
    </w:body></w:document>"#;

    assert_eq!(
        texts(xml),
        ["visible zero", "visible false", "real content"]
    );
}

#[test]
fn fixture_paragraph_id_contract_is_explicit() {
    for fixture in [
        BASIC_XML,
        COMMENT_XML,
        HIDDEN_XML,
        PLACEHOLDER_XML,
        TEXTBOX_XML,
    ] {
        let projection = project_document_xml(fixture, allocate).unwrap();
        let ids = projection
            .paragraphs
            .iter()
            .filter_map(|paragraph| paragraph.package_paragraph_id)
            .collect::<Vec<_>>();
        assert!(
            projection
                .paragraphs
                .iter()
                .all(|paragraph| paragraph.package_paragraph_id.is_some())
        );
        assert_eq!(ids.iter().copied().collect::<HashSet<_>>().len(), ids.len());
    }
    let missing = project_document_xml(MISSING_ID_XML, allocate).unwrap();
    assert!(
        missing
            .paragraphs
            .iter()
            .all(|paragraph| paragraph.package_paragraph_id.is_none())
    );
}

#[test]
fn resolves_predefined_and_numeric_xml_character_references() {
    let projection = project_document_xml(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Legal &amp; &#9746; &#x1F600; &lt; text</w:t></w:r></w:p></w:body></w:document>"#,
        allocate,
    )
    .unwrap();

    assert_eq!(projection.paragraphs[0].text, "Legal & ☒ 😀 < text");
}

#[test]
fn applies_xml_space_after_collecting_each_text_element() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>  A &amp; B  </w:t></w:r></w:p>
      <w:p><w:r><w:t xml:space="preserve">  C &amp; D  </w:t></w:r></w:p>
      <w:p><w:del><w:r><w:delText>  old &amp; stale  </w:delText></w:r></w:del></w:p>
    </w:body></w:document>"#;

    assert_eq!(
        texts_in_view(xml, RevisionView::Original),
        ["A & B", "  C & D  ", "old & stale"]
    );
}

#[test]
fn resolves_wordprocessing_namespaces_independently_of_prefix() {
    for namespace in [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let xml = format!(
            r#"<x:document xmlns:x="{namespace}" xmlns:id="http://schemas.microsoft.com/office/word/2010/wordml"><x:body><x:p id:paraId="00000001"><x:r><x:t>Alias</x:t></x:r></x:p></x:body></x:document>"#,
        );
        let projection = project_document_xml(xml.as_bytes(), allocate).unwrap();
        assert_eq!(projection.paragraphs[0].text, "Alias");
        assert_eq!(
            projection.paragraphs[0].package_paragraph_id,
            PackageParagraphId::parse("00000001")
        );
    }

    let spoofed = br#"<w:document xmlns:w="urn:not-wordprocessingml"><w:body><w:p><w:r><w:t>Wrong namespace</w:t></w:r></w:p></w:body></w:document>"#;
    assert_eq!(
        project_document_xml(spoofed, allocate),
        Err(ProjectionError::MissingDocumentBody)
    );

    let aliased_package = package(
        &[
            (
                "word/document.xml",
                br#"<d:document xmlns:d="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><d:body><d:p><d:pPr><d:pStyle d:val="Body"/></d:pPr><d:r><d:t>Styled</d:t></d:r></d:p></d:body></d:document>"#,
            ),
            (
                "word/styles.xml",
                br#"<s:styles xmlns:s="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><s:style s:type="paragraph" s:styleId="Body"><s:pPr><s:ind s:left="720"/></s:pPr></s:style></s:styles>"#,
            ),
        ],
        CompressionMethod::Deflated,
    );
    let aliased = project_docx(&aliased_package, DocxLimits::default(), allocate).unwrap();
    let StructuralFactSet::Known(indentation) = aliased.structural_facts.indentation else {
        panic!("aliased styles should resolve");
    };
    assert_eq!(indentation[0].value.left_twips, Some(720));
}

#[test]
fn selects_exactly_one_supported_alternate_content_branch() {
    let supported_after_unknown = br#"
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:future="urn:future-wordprocessing">
        <w:body><mc:AlternateContent>
          <mc:Choice Requires="future"><w:p><w:r><w:t>unknown</w:t></w:r></w:p></mc:Choice>
          <mc:Choice Requires="w"><w:p><w:r><w:t>selected</w:t></w:r></w:p></mc:Choice>
          <mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback>
        </mc:AlternateContent></w:body>
      </w:document>
    "#;
    assert_eq!(texts(supported_after_unknown), ["selected"]);

    let fallback = br#"
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:future="urn:future-wordprocessing">
        <w:body><mc:AlternateContent>
          <mc:Choice Requires="future"><w:p><w:r><w:t>unknown</w:t></w:r></w:p></mc:Choice>
          <mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback>
        </mc:AlternateContent></w:body>
      </w:document>
    "#;
    assert_eq!(texts(fallback), ["fallback"]);

    let transparent = br#"
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <w:body><w:tbl><mc:AlternateContent>
          <mc:Choice Requires="w"><w:tr><w:tc><w:p><w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="w"><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></mc:Choice>
              <mc:Fallback><w:t>Wrong</w:t></mc:Fallback>
            </mc:AlternateContent>
          </w:r></w:p></w:tc></w:tr></mc:Choice>
          <mc:Fallback><w:tr><w:tc><w:p><w:r><w:t>Wrong</w:t></w:r></w:p></w:tc></w:tr></mc:Fallback>
        </mc:AlternateContent></w:tbl></w:body>
      </w:document>
    "#;
    let projection = project_document_xml(transparent, allocate)
        .expect("selected compatibility branches should be semantically transparent");
    assert_eq!(projection.paragraphs.len(), 1);
    assert_eq!(projection.paragraphs[0].text, "Bold");
    assert_eq!(
        projection.paragraphs[0].structure,
        Some(ParagraphStructure {
            table_ordinal: 0,
            row: 0,
            column: 0,
        })
    );
    assert_eq!(
        projection.paragraphs[0].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 4,
            style: TextStyle::Bold,
        }]
    );
}

#[test]
fn materializes_typed_text_controls_for_each_consumer_view() {
    let xml = br#"<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p><x:r><x:t>A</x:t><x:tab/><x:br/><x:br x:type="page"/><x:br x:type="column"/><x:cr/><x:footnoteReference/><x:softHyphen/><x:noBreakHyphen/><x:t>Z</x:t></x:r></x:p></x:body></x:document>"#;

    let word = project_document_xml(xml, allocate).unwrap();
    assert_eq!(
        word.paragraphs[0].text,
        "A\t\u{000b}\u{000c}\u{000e}\r\u{0002}\u{001f}\u{001e}Z"
    );

    let plain = project_document_xml_with_options(
        xml,
        ProjectionOptions {
            text_materialization: TextMaterialization::ReadablePlainText,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .unwrap();
    assert_eq!(plain.paragraphs[0].text, "A\t\n\n\n\n\u{00ad}\u{2011}Z");
}

#[test]
fn projects_direct_bold_utf16_offsets_and_table_facts() {
    let xml = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="01A2B3C4"><w:r><w:t>😀 </w:t></w:r></w:p>
          <w:tbl><w:tr>
            <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Left</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc>
          </w:tr></w:tbl>
          <w:p/>
        </w:body>
      </w:document>
    "#;
    let projection =
        project_document_xml(xml.as_bytes(), allocate).expect("synthetic OOXML should project");

    assert_eq!(projection.paragraphs[0].id.as_str(), "internal-0");
    assert_eq!(
        projection.paragraphs[0].package_paragraph_id,
        PackageParagraphId::parse("01A2B3C4")
    );
    assert_eq!(projection.paragraphs[0].ordinal, 0);
    assert_eq!(
        projection.paragraphs[1].formatting,
        vec![TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 4,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[1].structure,
        Some(ParagraphStructure {
            table_ordinal: 0,
            row: 0,
            column: 0,
        })
    );
    assert_eq!(
        projection.paragraphs[2].structure,
        Some(ParagraphStructure {
            table_ordinal: 0,
            row: 0,
            column: 1,
        })
    );
    assert_eq!(projection.paragraphs[3].text, "");
}

#[test]
fn preserves_table_coordinates_through_row_and_cell_wrappers() {
    let xml = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:tbl>
          <w:customXml w:element="row"><w:tr>
            <w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>
            <w:customXml w:element="cell"><w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc></w:customXml>
          </w:tr></w:customXml>
          <w:sdt><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>Next</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>
        </w:tbl></w:body>
      </w:document>
    "#;
    let projection = project_document_xml(xml, allocate).unwrap();

    assert_eq!(
        projection
            .paragraphs
            .iter()
            .map(|paragraph| (paragraph.text.as_str(), paragraph.structure.clone()))
            .collect::<Vec<_>>(),
        [
            (
                "Left",
                Some(ParagraphStructure {
                    table_ordinal: 0,
                    row: 0,
                    column: 0,
                }),
            ),
            (
                "Right",
                Some(ParagraphStructure {
                    table_ordinal: 0,
                    row: 0,
                    column: 1,
                }),
            ),
            (
                "Next",
                Some(ParagraphStructure {
                    table_ordinal: 0,
                    row: 1,
                    column: 0,
                }),
            ),
        ]
    );
}

#[test]
fn projects_only_semantic_direct_highlights_with_unicode_offsets() {
    let xml = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p>
          <w:r><w:rPr><w:highlight w:val="cyan"/></w:rPr><w:t>😀</w:t></w:r>
          <w:r><w:rPr><w:highlight w:val="lightGray"/></w:rPr><w:t>Muted</w:t></w:r>
          <w:r><w:rPr><w:highlight w:val="darkGray"/></w:rPr><w:t>Review</w:t></w:r>
          <w:r><w:rPr><w:shd w:val="clear" w:fill="FFFF00"/></w:rPr><w:t>Shaded</w:t></w:r>
          <w:r><w:rPr><w:b/><w:highlight w:val="yellow"/></w:rPr><w:t>Both</w:t></w:r>
          <w:r><w:rPr><w:highlight w:val="none"/></w:rPr><w:t>None</w:t></w:r>
          <w:r><w:rPr><w:highlight w:val="unknown"/></w:rPr><w:t>Unknown</w:t></w:r>
        </w:p></w:body>
      </w:document>
    "#;
    let projection =
        project_document_xml(xml.as_bytes(), allocate).expect("synthetic OOXML should project");

    assert_eq!(
        projection.paragraphs[0].text,
        "😀MutedReviewShadedBothNoneUnknown"
    );
    assert_eq!(
        projection.paragraphs[0].formatting,
        vec![
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 13,
                style: TextStyle::Highlight,
            },
            TextFormattingSpan {
                start_utf16: 19,
                end_utf16: 23,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 19,
                end_utf16: 23,
                style: TextStyle::Highlight,
            },
        ]
    );
}

#[test]
fn semantic_highlight_color_contract_is_exhaustive() {
    let contract = include_str!("../fixtures/ooxml-highlight-colors.tsv");
    let cases = contract
        .lines()
        .map(|line| {
            let (color, classification) = line.split_once('\t').expect("valid contract row");
            (color, classification == "semantic")
        })
        .collect::<Vec<_>>();

    assert_eq!(cases.len(), 17);
    for (color, expected) in cases {
        assert_eq!(
            stella_docx_kernel::is_semantic_highlight_color(color),
            expected,
            "{color}"
        );
    }
    assert!(!stella_docx_kernel::is_semantic_highlight_color("unknown"));
}

#[test]
fn keeps_package_and_application_identity_distinct() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
      <w:p w14:paraId="01A2B3C4"><w:r><w:t>One</w:t></w:r></w:p>
      <w:p w14:paraId="not-a-package-id"><w:r><w:t>Two</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let projection = project_document_xml(xml, |facts| {
        let package = facts.package_paragraph_id.map_or_else(
            || "none".to_owned(),
            |id| format!("package-{:08X}", id.value()),
        );
        InternalParagraphId::new(format!("app-{package}-{}", facts.ordinal))
    })
    .expect("identity allocation should succeed");

    assert_eq!(
        projection.paragraphs[0].id.as_str(),
        "app-package-01A2B3C4-0"
    );
    assert_eq!(projection.paragraphs[1].id.as_str(), "app-none-1");
    assert_eq!(projection.paragraphs[1].package_paragraph_id, None);

    let duplicate = project_document_xml(xml, |_| InternalParagraphId::new("same"));
    assert_eq!(
        duplicate,
        Err(ProjectionError::DuplicateInternalParagraphId)
    );
}

#[test]
fn extracts_store_and_deflate_packages_before_projection() {
    for method in [CompressionMethod::Stored, CompressionMethod::Deflated] {
        let package = package(&[("word/document.xml", BASIC_XML)], method);
        assert_eq!(
            extract_document_xml(&package, DocxLimits::default())
                .expect("main part should extract"),
            BASIC_XML
        );
        assert_eq!(
            project_docx(&package, DocxLimits::default(), allocate)
                .expect("DOCX should project")
                .paragraphs
                .len(),
            4
        );
    }
}

#[test]
fn resolves_the_main_document_from_package_relationships() {
    let custom_document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Relationship target</w:t></w:r></w:p></w:body>
      </w:document>
    "#;
    let relationships = br#"
      <r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
        <r:Relationship Id="metadata" Type="urn:unrelated" Target="metadata.xml"/>
        <r:Relationship
          Id="document"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="./custom/main.xml"/>
      </r:Relationships>
    "#;
    let conventional_document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Wrong conventional part</w:t></w:r></w:p></w:body>
      </w:document>
    "#;
    let package = package(
        &[
            ("_rels/.rels", relationships),
            ("custom/main.xml", custom_document),
            ("word/document.xml", conventional_document),
        ],
        CompressionMethod::Deflated,
    );

    assert_eq!(
        extract_document_xml(&package, DocxLimits::default())
            .expect("relationship-resolved main part should extract"),
        custom_document
    );
    assert_eq!(
        project_docx(&package, DocxLimits::default(), allocate)
            .expect("relationship-resolved main part should project")
            .paragraphs
            .into_iter()
            .map(|paragraph| paragraph.text)
            .collect::<Vec<_>>(),
        ["Relationship target"]
    );
}

#[test]
fn rejects_ambiguous_or_unbounded_archives() {
    let missing = package(&[("unrelated.xml", BASIC_XML)], CompressionMethod::Stored);
    assert_eq!(
        extract_document_xml(&missing, DocxLimits::default()),
        Err(ProjectionError::MissingDocumentXml)
    );

    let mut duplicate = package(
        &[
            ("word/document.xml", BASIC_XML),
            ("word/document.xm_", BASIC_XML),
        ],
        CompressionMethod::Stored,
    );
    replace_all_equal_length(&mut duplicate, b"word/document.xm_", b"word/document.xml");
    assert_eq!(
        extract_document_xml(&duplicate, DocxLimits::default()),
        Err(ProjectionError::DuplicateDocumentXml)
    );

    let deflated = package(
        &[("word/document.xml", BASIC_XML)],
        CompressionMethod::Deflated,
    );
    let ratio_limit = DocxLimits {
        maximum_compression_ratio: 1,
        ..DocxLimits::default()
    };
    assert_eq!(
        extract_document_xml(&deflated, ratio_limit),
        Err(ProjectionError::SuspiciousCompressionRatio)
    );

    let tiny_xml_limit = DocxLimits {
        maximum_document_xml_bytes: 8,
        ..DocxLimits::default()
    };
    assert_eq!(
        extract_document_xml(&deflated, tiny_xml_limit),
        Err(ProjectionError::DocumentXmlTooLarge)
    );

    let tiny_archive_limit = DocxLimits {
        maximum_archive_bytes: 8,
        ..DocxLimits::default()
    };
    assert_eq!(
        extract_document_xml(&deflated, tiny_archive_limit),
        Err(ProjectionError::ArchiveTooLarge)
    );
}

#[test]
fn rejects_encryption_unsupported_compression_and_crc_drift() {
    let stored = package(
        &[("word/document.xml", BASIC_XML)],
        CompressionMethod::Stored,
    );

    let mut encrypted = stored.clone();
    mutate_zip_headers(&mut encrypted, |flags, _method| *flags |= 1);
    assert_eq!(
        extract_document_xml(&encrypted, DocxLimits::default()),
        Err(ProjectionError::EncryptedDocumentXml)
    );

    let mut unsupported = stored.clone();
    mutate_zip_headers(&mut unsupported, |_flags, method| *method = 9);
    assert_eq!(
        extract_document_xml(&unsupported, DocxLimits::default()),
        Err(ProjectionError::UnsupportedCompression(9))
    );

    let mut corrupt = stored;
    let xml_offset = find_bytes(&corrupt, BASIC_XML).expect("stored XML should be literal");
    corrupt[xml_offset] ^= 1;
    assert_eq!(
        extract_document_xml(&corrupt, DocxLimits::default()),
        Err(ProjectionError::DocumentXmlIntegrity)
    );

    let styles =
        br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let mut corrupt_styles = package(
        &[
            ("word/document.xml", BASIC_XML),
            ("word/styles.xml", styles),
        ],
        CompressionMethod::Stored,
    );
    let styles_offset =
        find_bytes(&corrupt_styles, styles).expect("stored styles should be literal");
    corrupt_styles[styles_offset] ^= 1;
    assert_eq!(
        extract_document_parts(&corrupt_styles, DocxLimits::default()),
        Err(ProjectionError::StylesXmlIntegrity)
    );
}

#[test]
fn rejects_multiple_document_bodies() {
    let xml =
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body><w:body><w:p/></w:body></w:document>"#;

    assert_eq!(
        project_document_xml(xml, allocate),
        Err(ProjectionError::InvalidDocumentXml)
    );
}

#[test]
fn rejects_paragraph_counts_above_the_configured_limit() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:p/></w:body></w:document>"#;
    let package = package(&[("word/document.xml", xml)], CompressionMethod::Deflated);
    let limits = DocxLimits {
        maximum_paragraphs: 1,
        ..DocxLimits::default()
    };

    assert_eq!(
        project_docx(&package, limits, allocate),
        Err(ProjectionError::TooManyParagraphs)
    );
}

#[test]
fn rejects_structural_fact_expansion_above_the_configured_limit() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:bookmarkStart w:id="1" w:name="first"/><w:bookmarkStart w:id="2" w:name="second"/><w:r><w:t>A</w:t></w:r></w:p>
      <w:p><w:r><w:t>B</w:t></w:r><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="1"/></w:p>
    </w:body></w:document>"#;
    let package = package(&[("word/document.xml", xml)], CompressionMethod::Deflated);
    let limits = DocxLimits {
        maximum_structural_facts: 3,
        ..DocxLimits::default()
    };

    assert_eq!(
        project_docx(&package, limits, allocate),
        Err(ProjectionError::TooManyStructuralFacts)
    );
}

#[test]
fn resolves_inherited_indentation_and_numbering_without_partial_facts() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="ListRoot"/></w:pPr><w:r><w:t>Root</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="ListChild"/></w:pPr><w:r><w:t>Child</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Normal"/><w:ind w:firstLine="240"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>
      </w:body></w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:pPrDefault><w:pPr><w:ind w:left="120"/></w:pPr></w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
        <w:style w:type="paragraph" w:styleId="ListRoot">
          <w:basedOn w:val="Normal"/>
          <w:pPr><w:ind w:left="720" w:hanging="360"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="ListChild">
          <w:basedOn w:val="ListRoot"/>
          <w:pPr><w:ind w:left="1440"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>
        </w:style>
      </w:styles>
    "#;
    let package = package(
        &[("word/document.xml", document), ("word/styles.xml", styles)],
        CompressionMethod::Deflated,
    );

    let projection = project_docx(&package, DocxLimits::default(), allocate)
        .expect("supported package structure should project");
    let StructuralFactSet::Known(indentation) = &projection.structural_facts.indentation else {
        panic!("resolved indentation must be known");
    };
    assert_eq!(indentation.len(), 3);
    assert_eq!(indentation[0].value.left_twips, Some(720));
    assert_eq!(indentation[0].value.hanging_twips, Some(360));
    assert_eq!(indentation[1].value.left_twips, Some(1440));
    assert_eq!(indentation[1].value.hanging_twips, Some(360));
    assert_eq!(indentation[2].value.left_twips, Some(120));
    assert_eq!(indentation[2].value.first_line_twips, Some(240));

    let StructuralFactSet::Known(numbering) = &projection.structural_facts.numbering_hierarchy
    else {
        panic!("resolved numbering must be known");
    };
    assert_eq!(numbering.len(), 2);
    assert_eq!(numbering[0].paragraph_ordinal, 0);
    assert_eq!(numbering[0].parent_paragraph_ordinal, None);
    assert_eq!(numbering[0].child_paragraph_ordinals, [1]);
    assert_eq!(numbering[1].parent_paragraph_ordinal, Some(0));
}

#[test]
fn document_part_only_does_not_claim_style_dependent_facts() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:body></w:document>"#;
    let projection = project_document_xml(xml, allocate).expect("document part should project");

    assert_eq!(
        projection.structural_facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::DocumentPartOnly)
    );
    assert_eq!(
        projection.structural_facts.numbering_hierarchy,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::DocumentPartOnly)
    );
    assert_eq!(
        projection.structural_facts.bookmarks,
        StructuralFactSet::Known(Vec::new())
    );
    assert_eq!(
        projection.structural_facts.internal_references,
        StructuralFactSet::Known(Vec::new())
    );
}

#[test]
fn projects_unicode_bookmark_and_internal_reference_boundaries() {
    let document = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:r><w:t>A</w:t></w:r><w:bookmarkStart w:id="9" w:name="Target"/><w:r><w:t>😀</w:t></w:r></w:p>
        <w:p><w:r><w:t>é</w:t></w:r><w:bookmarkEnd w:id="9"/><w:r><w:t>Z</w:t></w:r></w:p>
        <w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="Target"><w:r><w:t>here</w:t></w:r></w:hyperlink></w:p>
      </w:body></w:document>
    "#
    .replace('é', "e\u{301}");
    let package = package_with_minimal_styles(document.as_bytes());
    let projection = project_docx(&package, DocxLimits::default(), allocate)
        .expect("Unicode structure should project");

    let StructuralFactSet::Known(bookmarks) = &projection.structural_facts.bookmarks else {
        panic!("bookmark ranges must be known");
    };
    assert_eq!(bookmarks.len(), 2);
    assert_eq!(bookmarks[0].span.start_utf8, 1);
    assert_eq!(bookmarks[0].span.end_utf8, 5);
    assert_eq!(bookmarks[0].span.start_utf16, 1);
    assert_eq!(bookmarks[0].span.end_utf16, 3);
    assert_eq!(bookmarks[0].span.coverage, SpanCoverage::ContinuesAfter);
    assert_eq!(bookmarks[1].span.start_utf8, 0);
    assert_eq!(bookmarks[1].span.end_utf8, 3);
    assert_eq!(bookmarks[1].span.start_utf16, 0);
    assert_eq!(bookmarks[1].span.end_utf16, 2);
    assert_eq!(bookmarks[1].span.coverage, SpanCoverage::ContinuesBefore);

    let StructuralFactSet::Known(references) = &projection.structural_facts.internal_references
    else {
        panic!("hyperlink reference must be known");
    };
    assert_eq!(references.len(), 3);
    assert_eq!(references[0].paragraph_ordinal, 0);
    assert_eq!(references[0].role, InternalReferenceRole::Target);
    assert_eq!(references[1].paragraph_ordinal, 1);
    assert_eq!(references[1].role, InternalReferenceRole::Target);
    assert_eq!(references[2].paragraph_ordinal, 2);
    assert_eq!(references[2].role, InternalReferenceRole::Source);
    assert_eq!(references[2].span.start_utf8, 4);
    assert_eq!(references[2].span.start_utf16, 4);
    assert_eq!(references[2].span.start_utf8, references[2].span.end_utf8);
    assert_eq!(references[2].span.start_utf16, references[2].span.end_utf16);
}

#[test]
fn emits_each_internal_reference_target_once() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:bookmarkStart w:id="9" w:name="Target"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>
      <w:p><w:hyperlink w:anchor="Target"><w:r><w:t>First</w:t></w:r></w:hyperlink></w:p>
      <w:p><w:hyperlink w:anchor="Target"><w:r><w:t>Second</w:t></w:r></w:hyperlink></w:p>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .unwrap();
    let StructuralFactSet::Known(references) = projection.structural_facts.internal_references
    else {
        panic!("supported references should be known");
    };

    assert_eq!(references.len(), 3);
    assert_eq!(
        references
            .iter()
            .filter(|reference| reference.role == InternalReferenceRole::Target)
            .count(),
        1
    );
    assert_eq!(
        references
            .iter()
            .filter(|reference| reference.role == InternalReferenceRole::Source)
            .map(|reference| reference.paragraph_ordinal)
            .collect::<Vec<_>>(),
        [1, 2]
    );
}

#[test]
fn downgrades_whole_fact_families_for_unsupported_or_incomplete_constructs() {
    let incomplete_bookmark = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="1" w:name="Open"/><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>"#;
    let incomplete_projection = project_docx(
        &package_with_minimal_styles(incomplete_bookmark),
        DocxLimits::default(),
        allocate,
    )
    .expect("incomplete bookmark metadata should not abort paragraph projection");
    assert_eq!(
        incomplete_projection.structural_facts.bookmarks,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::IncompleteBookmarkRanges)
    );
    assert_eq!(
        incomplete_projection.structural_facts.internal_references,
        StructuralFactSet::Known(Vec::new()),
        "no reference source remains authoritative even when bookmarks are unavailable"
    );

    let field_reference = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF Target </w:instrText></w:r></w:p></w:body></w:document>"#;
    let field_projection = project_docx(
        &package_with_minimal_styles(field_reference),
        DocxLimits::default(),
        allocate,
    )
    .expect("unbalanced fields should preserve paragraph projection");
    assert_eq!(
        field_projection.structural_facts.internal_references,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedInternalReferences)
    );

    let skipped_level = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Orphan</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let skipped_level_projection = project_docx(
        &package_with_minimal_styles(skipped_level),
        DocxLimits::default(),
        allocate,
    )
    .expect("a nonzero first level should project as a root");
    assert_eq!(
        skipped_level_projection
            .structural_facts
            .numbering_hierarchy,
        StructuralFactSet::Known(Vec::new()),
        "a root without relationships is omitted from the compact hierarchy"
    );
    assert!(matches!(
        skipped_level_projection.structural_facts.indentation,
        StructuralFactSet::Known(_)
    ));

    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="CycleA"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>"#;
    let cyclic_styles = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="CycleA" w:default="1"><w:basedOn w:val="CycleB"/></w:style>
      <w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style>
    </w:styles>"#;
    let cyclic_style_projection = project_docx(
        &package(
            &[
                ("word/document.xml", document),
                ("word/styles.xml", cyclic_styles),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("unsupported style graphs should preserve paragraph projection");
    assert_eq!(
        cyclic_style_projection.structural_facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        cyclic_style_projection.structural_facts.numbering_hierarchy,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
}

#[test]
fn projects_standard_field_references_and_body_level_bookmark_boundaries() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:bookmarkStart w:id="4" w:name="Target"/><w:r><w:t>Target text</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="4"/>
      <w:p>
        <w:r><w:t xml:space="preserve">See </w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> REF </w:instrText><w:instrText>"Target" \h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>target text</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("standard fields and body boundaries should project");

    let StructuralFactSet::Known(bookmarks) = projection.structural_facts.bookmarks else {
        panic!("body-level bookmark end should remain complete");
    };
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].paragraph_ordinal, 0);
    assert_eq!(bookmarks[0].span.start_utf8, 0);
    assert_eq!(bookmarks[0].span.end_utf8, 11);

    let StructuralFactSet::Known(references) = projection.structural_facts.internal_references
    else {
        panic!("supported REF and unrelated PAGE fields should be complete");
    };
    assert_eq!(references.len(), 2);
    assert_eq!(references[0].role, InternalReferenceRole::Target);
    assert_eq!(references[1].role, InternalReferenceRole::Source);
    assert_eq!(references[1].paragraph_ordinal, 1);
    assert_eq!(references[1].span.start_utf8, 4);
}

#[test]
fn package_projection_is_deterministic_and_styles_extraction_is_bounded() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Stable</w:t></w:r></w:p></w:body></w:document>"#;
    let package_bytes = package_with_minimal_styles(document);
    let first = project_docx(&package_bytes, DocxLimits::default(), allocate).unwrap();
    let second = project_docx(&package_bytes, DocxLimits::default(), allocate).unwrap();
    assert_eq!(first, second);
    assert!(
        extract_document_parts(&package_bytes, DocxLimits::default())
            .unwrap()
            .styles_xml
            .is_some()
    );

    let limits = DocxLimits {
        maximum_styles_xml_bytes: 8,
        ..DocxLimits::default()
    };
    assert_eq!(
        extract_document_parts(&package_bytes, limits),
        Err(ProjectionError::StylesXmlTooLarge)
    );

    let equal_limits_package = package(
        &[
            ("word/document.xml", &b"<w:document/>"[..]),
            ("word/styles.xml", &b"<w:styles><w:style/></w:styles>"[..]),
        ],
        CompressionMethod::Stored,
    );
    let equal_limits = DocxLimits {
        maximum_document_xml_bytes: 16,
        maximum_styles_xml_bytes: 16,
        ..DocxLimits::default()
    };
    assert_eq!(
        extract_document_parts(&equal_limits_package, equal_limits),
        Err(ProjectionError::StylesXmlTooLarge)
    );
}

fn package_with_minimal_styles(document: &[u8]) -> Vec<u8> {
    package(
        &[
            ("word/document.xml", document),
            (
                "word/styles.xml",
                br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/></w:styles>"#,
            ),
        ],
        CompressionMethod::Deflated,
    )
}

fn package(entries: &[(&str, &[u8])], method: CompressionMethod) -> Vec<u8> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(method);
    for (name, contents) in entries {
        archive
            .start_file(*name, options)
            .expect("test ZIP entry should start");
        archive
            .write_all(contents)
            .expect("test ZIP entry should write");
    }
    archive
        .finish()
        .expect("test ZIP should finish")
        .into_inner()
}

fn mutate_zip_headers(bytes: &mut [u8], mut mutate: impl FnMut(&mut u16, &mut u16)) {
    let mut offset = 0;
    while offset + 12 <= bytes.len() {
        let (flags_offset, method_offset) = if bytes[offset..].starts_with(b"PK\x03\x04") {
            (offset + 6, offset + 8)
        } else if bytes[offset..].starts_with(b"PK\x01\x02") {
            (offset + 8, offset + 10)
        } else {
            offset += 1;
            continue;
        };
        let mut flags = u16::from_le_bytes([bytes[flags_offset], bytes[flags_offset + 1]]);
        let mut method = u16::from_le_bytes([bytes[method_offset], bytes[method_offset + 1]]);
        mutate(&mut flags, &mut method);
        bytes[flags_offset..flags_offset + 2].copy_from_slice(&flags.to_le_bytes());
        bytes[method_offset..method_offset + 2].copy_from_slice(&method.to_le_bytes());
        offset = method_offset + 2;
    }
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn replace_all_equal_length(haystack: &mut [u8], from: &[u8], to: &[u8]) {
    assert_eq!(
        from.len(),
        to.len(),
        "replacement must preserve ZIP offsets"
    );
    let mut offset = 0;
    while let Some(relative) = find_bytes(&haystack[offset..], from) {
        let start = offset + relative;
        haystack[start..start + to.len()].copy_from_slice(to);
        offset = start + to.len();
    }
}
