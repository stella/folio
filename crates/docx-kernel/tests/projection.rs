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
use std::fmt::Write as _;
use std::io::{Cursor, Write};

use stella_docx_kernel::ParagraphAlignmentValue as Align;
use stella_docx_kernel::{
    CommentContent, DocxLimits, FormattingProjectionStatus, FormattingUnknownReason,
    InternalParagraphId, InternalReferenceRole, PackageParagraphId, ParagraphAlignmentFact,
    ParagraphAlignmentSource, ParagraphIdentityFacts, ParagraphOutlineLevelFact,
    ParagraphStructure, ProjectionError, ProjectionOptions, ReviewDetail, ReviewFactLimits,
    ReviewFactSet, ReviewFactUnknownReason, ReviewPoint, ReviewSpan, RevisionContent,
    RevisionFactKind, RevisionProjectionStatus, RevisionUnsupportedReason, RevisionView,
    SpanCoverage, StructuralFactSet, StructuralFactUnknownReason, TextFormattingSpan,
    TextMaterialization, TextStyle, extract_document_parts, extract_document_xml,
    project_document_xml, project_document_xml_with_options, project_docx,
    project_docx_with_options, project_docx_with_review_facts,
};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const BASIC_XML: &[u8] = include_bytes!("../fixtures/basic.xml");
const COMMENT_XML: &[u8] = include_bytes!("../fixtures/comment.xml");
const HIDDEN_XML: &[u8] = include_bytes!("../fixtures/hidden.xml");
const MISSING_ID_XML: &[u8] = include_bytes!("../fixtures/missingParaId.xml");
const PLACEHOLDER_XML: &[u8] = include_bytes!("../fixtures/placeholder.xml");
const TEXTBOX_XML: &[u8] = include_bytes!("../fixtures/textbox.xml");
const MINIMAL_STYLES: &[u8] = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/></w:styles>"#;
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
            format!("test{SPECIAL_VISIBLE_TEXT}∮"),
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
fn projects_only_visible_office_math_text() {
    let xml = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:fake="urn:not-office-math">
        <w:body><w:p>
          <w:r><w:t>A</w:t></w:r>
          <m:oMath><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>x&amp;&#x1F600;</m:t></m:r></m:oMath>
          <fake:oMath><fake:r><fake:t>spoofed</fake:t></fake:r></fake:oMath>
          <m:t>detached</m:t>
          <w:r><w:t>Z</w:t></w:r>
        </w:p></w:body>
      </w:document>
    "#;

    assert_eq!(texts(xml.as_bytes()), ["Ax&😀Z"]);
}

#[test]
fn decodes_font_bound_word_symbols_without_trusting_spoofed_attributes() {
    let xml = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:fake="urn:not-wordprocessingml">
        <w:body><w:p><w:r>
          <w:sym w:font="Wingdings" w:char="F06C"/>
          <w:sym w:font="Courier New" w:char="F06C"/>
          <w:sym fake:font="Wingdings" w:font="Courier New" w:char="F06C"/>
        </w:r></w:p></w:body>
      </w:document>
    "#;

    assert_eq!(texts(xml), ["●\u{f06c}\u{f06c}"]);
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
            r#"<x:document xmlns:x="{namespace}" xmlns:id="http://schemas.microsoft.com/office/word/2010/wordml"><x:body><x:p id:paraId="00000001"><x:r><x:rPr><x:vertAlign x:val="superscript"/></x:rPr><x:t>Alias</x:t></x:r></x:p></x:body></x:document>"#,
        );
        let projection = project_document_xml(xml.as_bytes(), allocate).unwrap();
        assert_eq!(projection.paragraphs[0].text, "Alias");
        assert_eq!(
            projection.paragraphs[0].formatting,
            [TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 5,
                style: TextStyle::Superscript,
            }]
        );
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

    let partial_math_support_uses_fallback = br#"
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <w:body><mc:AlternateContent>
          <mc:Choice Requires="m"><w:p><w:r><w:t>unclaimed-math</w:t></w:r></w:p></mc:Choice>
          <mc:Fallback><w:p><w:r><w:t>math-fallback</w:t></w:r></w:p></mc:Fallback>
        </mc:AlternateContent></w:body>
      </w:document>
    "#;
    assert_eq!(texts(partial_math_support_uses_fallback), ["math-fallback"]);

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
                end_utf16: 2,
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
fn projects_only_direct_superscript_and_coalesces_adjacent_runs() {
    let xml = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p>
          <w:r><w:t>A</w:t></w:r>
          <w:r><w:rPr><w:b/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>😀)</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>x</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>B</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>C</w:t></w:r>
          <w:r><w:rPr><w:vertAlign/></w:rPr><w:t>D</w:t></w:r>
          <w:r><w:rPr><w:rPrChange w:id="1"><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:rPrChange></w:rPr><w:t>E</w:t></w:r>
        </w:p></w:body>
      </w:document>
    "#;
    let projection =
        project_document_xml(xml.as_bytes(), allocate).expect("synthetic OOXML should project");

    assert_eq!(projection.paragraphs[0].text, "A😀)xBCDE");
    assert_eq!(
        projection.paragraphs[0].formatting,
        vec![
            TextFormattingSpan {
                start_utf16: 1,
                end_utf16: 4,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 1,
                end_utf16: 4,
                style: TextStyle::Highlight,
            },
            TextFormattingSpan {
                start_utf16: 1,
                end_utf16: 5,
                style: TextStyle::Superscript,
            },
        ]
    );
}

#[test]
#[allow(clippy::too_many_lines)] // One table covers slot selection and exact UTF-16 spans.
fn selects_regular_and_complex_script_bold_from_the_effective_run_state() {
    let xml = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Aع&#x1F600;</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>Aع،B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b/><w:bCs w:val="false"/></w:rPr><w:t>Aع،B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b w:val="false"/><w:bCs/></w:rPr><w:t>Aع،B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b w:val="false"/><w:bCs/><w:cs/></w:rPr><w:t>forced</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b/><w:bCs w:val="false"/><w:cs/></w:rPr><w:t>forced</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/><w:cs w:val="0"/></w:rPr><w:t>عربي</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>&#x1F600;Aع،B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/><w:rtl/></w:rPr><w:t>Latin</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b/><w:bCs w:val="false"/><w:rtl/></w:rPr><w:t>Latin</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>&#x1820;</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>A&#x1E900;B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>A&#x11000;B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>A&#x301;ع&#x651;B</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>한</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>한</w:t></w:r></w:p>
        </w:body>
      </w:document>
    "#;
    let projection = project_document_xml(xml.as_bytes(), allocate).unwrap();

    assert_eq!(
        projection.paragraphs[0].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 1,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[1].formatting,
        [TextFormattingSpan {
            start_utf16: 1,
            end_utf16: 3,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[2].formatting,
        [
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 3,
                end_utf16: 4,
                style: TextStyle::Bold,
            },
        ]
    );
    assert_eq!(
        projection.paragraphs[3].formatting,
        [TextFormattingSpan {
            start_utf16: 1,
            end_utf16: 3,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[4].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 6,
            style: TextStyle::Bold,
        }]
    );
    assert!(projection.paragraphs[5].formatting.is_empty());
    assert_eq!(
        projection.paragraphs[6].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 4,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[7].formatting,
        [TextFormattingSpan {
            start_utf16: 3,
            end_utf16: 5,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.paragraphs[8].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 5,
            style: TextStyle::Bold,
        }]
    );
    assert!(projection.paragraphs[9].formatting.is_empty());
    assert_eq!(
        projection.paragraphs[10].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 1,
            style: TextStyle::Bold,
        }]
    );
    for paragraph in &projection.paragraphs[11..13] {
        assert_eq!(
            paragraph.formatting,
            [TextFormattingSpan {
                start_utf16: 1,
                end_utf16: 3,
                style: TextStyle::Bold,
            }]
        );
    }
    assert_eq!(
        projection.paragraphs[13].formatting,
        [TextFormattingSpan {
            start_utf16: 2,
            end_utf16: 4,
            style: TextStyle::Bold,
        }]
    );
    assert!(projection.paragraphs[14].formatting.is_empty());
    assert_eq!(
        projection.paragraphs[15].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 1,
            style: TextStyle::Bold,
        }]
    );
}

fn assert_effective_text_formatting_through_style_cascades(namespace: &str) {
    let document = format!(
        r#"
      <w:document xmlns:w="{namespace}">
        <w:body><w:p><w:pPr><w:pStyle w:val="Clause"/></w:pPr>
          <w:r><w:t>A</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>B</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="Raised"/></w:rPr><w:t>C</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="Raised"/><w:b w:val="0"/><w:highlight w:val="cyan"/><w:vertAlign w:val="baseline"/></w:rPr><w:t>D</w:t></w:r>
        </w:p></w:body>
      </w:document>
    "#
    );
    let styles = format!(
        r#"
      <w:styles xmlns:w="{namespace}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Clause" w:default="1"><w:basedOn w:val="Base"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:b w:val="off"/><w:highlight w:val="none"/><w:vertAlign w:val="baseline"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Raised"><w:basedOn w:val="Emphasis"/><w:rPr><w:b/><w:vertAlign w:val="superscript"/></w:rPr></w:style>
      </w:styles>
    "#
    );
    let projection = project_docx(
        &package(
            &[
                ("word/document.xml", document.as_bytes()),
                ("word/styles.xml", styles.as_bytes()),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .unwrap();

    assert_eq!(
        projection.paragraphs[0].text, "ABCD",
        "the namespace profile must not change visible text"
    );
    assert_eq!(
        projection.paragraphs[0].formatting,
        [
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 2,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Highlight,
            },
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Superscript,
            },
            TextFormattingSpan {
                start_utf16: 2,
                end_utf16: 3,
                style: TextStyle::Superscript,
            },
            TextFormattingSpan {
                start_utf16: 3,
                end_utf16: 4,
                style: TextStyle::Highlight,
            },
        ],
        "the namespace profile must not change effective formatting"
    );
}

#[test]
fn resolves_effective_text_formatting_through_transitional_style_cascades() {
    assert_effective_text_formatting_through_style_cascades(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    );
}

#[test]
fn resolves_effective_text_formatting_through_strict_style_cascades() {
    assert_effective_text_formatting_through_style_cascades(
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    );
}

#[test]
fn ignores_missing_and_cross_kind_based_on_targets_without_dropping_the_current_style() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="MissingParent"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="CrossKindParent"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>
      </w:body></w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="character" w:styleId="CharacterRoot"/>
        <w:style w:type="paragraph" w:styleId="MissingParent"><w:basedOn w:val="Absent"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="CrossKindParent"><w:basedOn w:val="CharacterRoot"/><w:rPr><w:b/></w:rPr></w:style>
      </w:styles>
    "#;
    let projection = project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("Word-ignored basedOn targets should preserve the current style");

    for paragraph in &projection.paragraphs {
        assert_eq!(
            paragraph.formatting,
            [TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Bold,
            }]
        );
    }
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Complete
    );
}

#[test]
fn resolves_markup_compatibility_inside_the_styles_part() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Compatible"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>"#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <w:style w:type="paragraph" w:styleId="Compatible"><w:rPr>
          <mc:AlternateContent>
            <mc:Choice Requires="w"><w:b/></mc:Choice>
            <mc:Fallback><w:b w:val="false"/></mc:Fallback>
          </mc:AlternateContent>
        </w:rPr></w:style>
      </w:styles>
    "#;
    let projection = project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("supported styles markup-compatibility content should project");

    assert_eq!(
        projection.paragraphs[0].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 1,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Complete
    );
}

#[test]
fn reports_numbering_and_math_formatting_hierarchies_as_incomplete() {
    let numbered_document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>"#;
    let numbering = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#;
    let numbered = project_docx(
        &package(
            &[
                ("word/document.xml", numbered_document),
                ("word/styles.xml", MINIMAL_STYLES),
                ("word/numbering.xml", numbering),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("numbering-level formatting should preserve best-known spans");
    assert!(numbered.paragraphs[0].formatting.is_empty());
    assert_eq!(
        numbered.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );

    let math_document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath><m:r><w:rPr><w:b/></w:rPr><m:t>A</m:t></m:r></m:oMath></w:p></w:body></w:document>"#;
    let math = project_docx(
        &package_with_minimal_styles(math_document),
        DocxLimits::default(),
        allocate,
    )
    .expect("math text should survive unsupported math formatting");
    assert_eq!(math.paragraphs[0].text, "A");
    assert!(math.paragraphs[0].formatting.is_empty());
    assert_eq!(
        math.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
}

#[test]
fn resolves_toggle_levels_and_ignores_historical_style_snapshots() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:pPr><w:pStyle w:val="BoldParagraph"/></w:pPr>
            <w:r><w:t>A</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="AlsoBold"/></w:rPr><w:t>B</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="ExplicitFalse"/></w:rPr><w:t>C</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="AlsoBold"/><w:b w:val="false"/></w:rPr><w:t>D</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="AlsoBold"/><w:b/></w:rPr><w:t>E</w:t></w:r>
          </w:p>
          <w:p><w:pPr><w:pStyle w:val="Historical"/></w:pPr><w:r><w:t>F</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:rStyle w:val="DefaultParagraphFont"/></w:rPr><w:t>G</w:t></w:r></w:p>
          <w:p><w:pPr><w:pStyle w:val="DerivedTrue"/></w:pPr><w:r><w:t>H</w:t></w:r></w:p>
          <w:p><w:pPr><w:pStyle w:val="DerivedFalse"/></w:pPr><w:r><w:t>I</w:t></w:r></w:p>
        </w:body>
      </w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
        <w:style w:type="paragraph" w:styleId="BoldParagraph"><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="AlsoBold"><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="ExplicitFalse"><w:rPr><w:b w:val="false"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Historical"><w:rPr><w:rPrChange w:id="1"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr></w:style>
        <w:style w:type="character" w:styleId="DefaultParagraphFont"><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="BoldBase"><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="MiddleFalse"><w:basedOn w:val="BoldBase"/><w:rPr><w:b w:val="false"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="DerivedTrue"><w:basedOn w:val="MiddleFalse"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="DerivedFalse"><w:basedOn w:val="MiddleFalse"/><w:rPr><w:b w:val="false"/></w:rPr></w:style>
      </w:styles>
    "#;
    let projection = project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("toggle fixtures should project");

    assert_eq!(
        projection.paragraphs[0].formatting,
        [
            TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 2,
                end_utf16: 3,
                style: TextStyle::Bold,
            },
            TextFormattingSpan {
                start_utf16: 4,
                end_utf16: 5,
                style: TextStyle::Bold,
            },
        ]
    );
    assert!(projection.paragraphs[1].formatting.is_empty());
    assert!(projection.paragraphs[2].formatting.is_empty());
    assert!(projection.paragraphs[3].formatting.is_empty());
    assert_eq!(
        projection.paragraphs[4].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 1,
            style: TextStyle::Bold,
        }]
    );
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Complete
    );
}

#[test]
fn ignores_style_identifiers_beyond_the_word_limit() {
    let overlong_id = "x".repeat(254);
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
          <w:p><w:pPr><w:pStyle w:val="{overlong_id}"/></w:pPr>
            <w:r><w:rPr><w:rStyle w:val="{overlong_id}"/></w:rPr><w:t>Plain</w:t></w:r>
          </w:p>
        </w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
          <w:style w:type="paragraph" w:styleId="{overlong_id}"><w:rPr><w:b/></w:rPr></w:style>
          <w:style w:type="character" w:styleId="{overlong_id}"><w:rPr><w:b/></w:rPr></w:style>
        </w:styles>"#
    );
    let projection = project_docx(
        &package(
            &[
                ("word/document.xml", document.as_bytes()),
                ("word/styles.xml", styles.as_bytes()),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("Word-ignored style identifiers should not invalidate the package");

    assert_eq!(projection.paragraphs[0].style_id, None);
    assert!(projection.paragraphs[0].formatting.is_empty());
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Complete
    );
}

#[test]
fn style_toggle_chains_preserve_false_and_direct_values_remain_absolute() {
    let document = r#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:pPr><w:pStyle w:val="Cancel"/></w:pPr>
            <w:r><w:t>ع</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="Restore"/></w:rPr><w:t>ب</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="Restore"/><w:bCs w:val="false"/></w:rPr><w:t>ج</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Latin</w:t></w:r></w:p>
          <w:p><w:pPr><w:pStyle w:val="Forced"/></w:pPr>
            <w:r><w:t>Latin</w:t></w:r>
            <w:r><w:rPr><w:rtl w:val="false"/></w:rPr><w:t>plain</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:rPrDefault><w:rPr><w:bCs/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="CancelBase"><w:rPr><w:bCs/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="CancelMiddle"><w:basedOn w:val="CancelBase"/><w:rPr><w:bCs w:val="0"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Cancel"><w:basedOn w:val="CancelMiddle"/><w:rPr><w:bCs/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Forced"><w:rPr><w:bCs/><w:rtl/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="RestoreBase"><w:rPr><w:bCs/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="RestoreMiddle"><w:basedOn w:val="RestoreBase"/><w:rPr><w:bCs w:val="false"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Restore"><w:basedOn w:val="RestoreMiddle"/><w:rPr><w:bCs/><w:cs/></w:rPr></w:style>
      </w:styles>
    "#;
    let projection = project_docx(
        &package(
            &[
                ("word/document.xml", document.as_bytes()),
                ("word/styles.xml", styles),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .unwrap();

    assert_eq!(projection.paragraphs[0].text, "عبج");
    assert_eq!(
        projection.paragraphs[0].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 2,
            style: TextStyle::Bold,
        }]
    );
    assert!(projection.paragraphs[1].formatting.is_empty());
    assert!(projection.paragraphs[2].formatting.is_empty());
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
fn resolves_styles_and_numbering_from_document_relationships() {
    for (package_namespace, office_namespace, word_namespace) in [
        (
            "http://schemas.openxmlformats.org/package/2006/relationships",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ),
        (
            "http://purl.oclc.org/ooxml/package/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ),
    ] {
        let root_relationships = format!(
            r#"<Relationships xmlns="{package_namespace}"><Relationship Type="{office_namespace}/officeDocument" Target="/custom/main.xml"/></Relationships>"#
        );
        let document_relationships = format!(
            r#"<Relationships xmlns="{package_namespace}"><Relationship Type="{office_namespace}/styles" Target="../shared/./discarded/../styles.xml"/><Relationship Type="{office_namespace}/numbering" Target="/lists/./nested/../numbering.xml" TargetMode="Internal"/></Relationships>"#
        );
        let document = format!(
            r#"<w:document xmlns:w="{word_namespace}"><w:body><w:p><w:pPr><w:pStyle w:val="List"/></w:pPr><w:r><w:t>Related</w:t></w:r></w:p></w:body></w:document>"#
        );
        let styles = format!(
            r#"<w:styles xmlns:w="{word_namespace}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>"#
        );
        let numbering = format!(
            r#"<w:numbering xmlns:w="{word_namespace}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num></w:numbering>"#
        );
        let package_bytes = package(
            &[
                ("_rels/.rels", root_relationships.as_bytes()),
                ("custom/main.xml", document.as_bytes()),
                (
                    "custom/_rels/main.xml.rels",
                    document_relationships.as_bytes(),
                ),
                ("shared/styles.xml", styles.as_bytes()),
                ("lists/numbering.xml", numbering.as_bytes()),
            ],
            CompressionMethod::Deflated,
        );

        let parts = extract_document_parts(&package_bytes, DocxLimits::default()).unwrap();
        assert_eq!(parts.styles_xml.as_deref(), Some(styles.as_bytes()));
        assert_eq!(parts.numbering_xml.as_deref(), Some(numbering.as_bytes()));
        let projection = project_docx(&package_bytes, DocxLimits::default(), allocate).unwrap();
        assert_eq!(projection.paragraphs[0].text, "Related");
        assert_eq!(projection.paragraphs[0].formatting.len(), 1);
        let StructuralFactSet::Known(indentation) = projection.structural_facts.indentation else {
            panic!("relationship-selected numbering should resolve indentation");
        };
        assert_eq!(indentation[0].value.left_twips, Some(720));
        assert_eq!(indentation[0].value.hanging_twips, Some(360));
    }
}

#[test]
fn ignores_unrelated_optional_parts_and_rejects_missing_relationship_targets() {
    let empty_relationships =
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Unrelated parts</w:t></w:r></w:p></w:body></w:document>"#;
    let unreferenced = package(
        &[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", empty_relationships),
            ("word/styles.xml", MINIMAL_STYLES),
            ("word/numbering.xml", b"<numbering/>"),
        ],
        CompressionMethod::Deflated,
    );
    let parts = extract_document_parts(&unreferenced, DocxLimits::default()).unwrap();
    assert_eq!(parts.styles_xml, None);
    assert_eq!(parts.numbering_xml, None);
    let projection = project_docx(&unreferenced, DocxLimits::default(), allocate).unwrap();
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::StylesPartUnavailable)
    );
    assert_eq!(
        projection.structural_facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::StylesPartUnavailable)
    );

    let missing_styles_relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="missing.xml"/></Relationships>"#;
    let missing = package(
        &[
            ("word/document.xml", BASIC_XML),
            ("word/_rels/document.xml.rels", missing_styles_relationships),
        ],
        CompressionMethod::Deflated,
    );
    assert_eq!(
        extract_document_parts(&missing, DocxLimits::default()),
        Err(ProjectionError::InvalidStylesXmlEntry)
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
fn counts_outline_levels_against_the_shared_structural_fact_budget() {
    let document: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
      <w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let package = package_with_minimal_styles(document);
    let constrained = DocxLimits {
        maximum_structural_facts: 1,
        ..DocxLimits::default()
    };

    assert_eq!(
        project_docx(&package, constrained, allocate),
        Err(ProjectionError::TooManyStructuralFacts)
    );
    let exact = DocxLimits {
        maximum_structural_facts: 2,
        ..DocxLimits::default()
    };
    assert!(project_docx(&package, exact, allocate).is_ok());
}

#[test]
fn rejects_outline_levels_outside_the_ooxml_range() {
    let direct = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:outlineLvl w:val="10"/></w:pPr></w:p></w:body></w:document>"#;
    assert_eq!(
        project_document_xml(direct, allocate),
        Err(ProjectionError::InvalidDocumentXml)
    );

    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>"#;
    let styles = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading"><w:pPr><w:outlineLvl w:val="10"/></w:pPr></w:style></w:styles>"#;
    let projection = project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("an unsupported optional styles part should preserve direct document facts");
    assert_eq!(
        projection.paragraphs[0].style_id.as_deref(),
        Some("Heading")
    );
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        projection.structural_facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
}

#[test]
fn resolves_inherited_indentation_and_numbering_without_partial_facts() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="ListRoot"/></w:pPr><w:r><w:t>Root</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="ListChild"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Child</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Normal"/><w:ind w:firstLine="240"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>
        <w:p><w:r><w:t>Default</w:t></w:r></w:p>
      </w:body></w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:pPrDefault><w:pPr><w:ind w:left="120"/><w:outlineLvl w:val="7"/></w:pPr></w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr><w:outlineLvl w:val="4"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="ListRoot">
          <w:basedOn w:val="Normal"/>
          <w:pPr><w:ind w:left="720" w:hanging="360"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="ListChild">
          <w:basedOn w:val="ListRoot"/>
          <w:pPr><w:ind w:left="1440"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>
        </w:style>
      </w:styles>
    "#;
    let numbering_xml = br#"
      <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"/><w:lvl w:ilvl="1"/></w:abstractNum>
        <w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>
      </w:numbering>
    "#;
    let package = package(
        &[
            ("word/document.xml", document),
            ("word/styles.xml", styles),
            ("word/numbering.xml", numbering_xml),
        ],
        CompressionMethod::Deflated,
    );

    let projection = project_docx(&package, DocxLimits::default(), allocate)
        .expect("supported package structure should project");
    assert_eq!(
        projection
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.style_id.as_deref())
            .collect::<Vec<_>>(),
        [Some("ListRoot"), Some("ListChild"), Some("Normal"), None]
    );
    let StructuralFactSet::Known(indentation) = &projection.structural_facts.indentation else {
        panic!("resolved indentation must be known");
    };
    assert_eq!(indentation.len(), 4);
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

    let StructuralFactSet::Known(outline_levels) = &projection.structural_facts.outline_levels
    else {
        panic!("resolved outline levels must be known");
    };
    assert_eq!(
        outline_levels
            .iter()
            .map(|fact| (fact.paragraph_ordinal, fact.outline_level))
            .collect::<Vec<_>>(),
        [(0, 1), (1, 2), (2, 4), (3, 4)]
    );
}

#[test]
fn resolves_numbering_level_indentation_through_the_ooxml_cascade() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="StyledList"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="StyledList"/><w:ind w:left="2000" w:firstLine="300"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Base"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr><w:ind w:firstLine="0"/></w:pPr><w:r><w:t>Level</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="6"/></w:numPr></w:pPr><w:r><w:t>Override</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Linked</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Base"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr><w:ind w:hanging="0"/></w:pPr><w:r><w:t>Neutral hanging</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Base"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr><w:ind w:hanging="180"/></w:pPr><w:r><w:t>Direct hanging</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="StyledList"/><w:numPr><w:numId w:val="0"/></w:numPr><w:ind w:left="357"/></w:pPr><w:r><w:t>Cancelled with direct indent</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="StyledList"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Cancelled without direct indent</w:t></w:r></w:p>
      </w:body></w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:pPrDefault><w:pPr><w:ind w:right="100"/></w:pPr></w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:ind w:left="1000"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="StyledList"><w:basedOn w:val="Base"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr></w:style>
      </w:styles>
    "#;
    let numbering = br#"
      <n:numbering xmlns:n="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <n:abstractNum n:abstractNumId="30"><n:styleLink n:val="Shared"/><n:lvl n:ilvl="1"><n:pPr><n:ind n:left="1417" n:right="200" n:hanging="793"/></n:pPr></n:lvl></n:abstractNum>
        <n:abstractNum n:abstractNumId="31"><n:numStyleLink n:val="Shared"/></n:abstractNum>
        <n:num n:numId="5"><n:abstractNumId n:val="30"/></n:num>
        <n:num n:numId="6"><n:abstractNumId n:val="30"/><n:lvlOverride n:ilvl="1"><n:lvl n:ilvl="1"><n:pPr><n:ind n:left="900" n:firstLine="120"/></n:pPr></n:lvl></n:lvlOverride></n:num>
        <n:num n:numId="7"><n:abstractNumId n:val="31"/></n:num>
      </n:numbering>
    "#;
    let projection = project_docx(
        &package(
            &[
                ("word/document.xml", document),
                ("word/styles.xml", styles),
                ("word/numbering.xml", numbering),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("supported numbering cascade should project");
    let StructuralFactSet::Known(indentation) = projection.structural_facts.indentation else {
        panic!("numbering-derived indentation should be known");
    };
    assert_eq!(indentation.len(), 8);
    assert_eq!(indentation[0].value.left_twips, Some(1000));
    assert_eq!(indentation[0].value.right_twips, Some(200));
    assert_eq!(indentation[0].value.hanging_twips, Some(793));
    assert_eq!(indentation[1].value.left_twips, Some(2000));
    assert_eq!(indentation[1].value.first_line_twips, Some(300));
    assert_eq!(indentation[1].value.hanging_twips, None);
    assert_eq!(indentation[2].value.left_twips, Some(1417));
    assert_eq!(indentation[2].value.hanging_twips, Some(793));
    assert_eq!(indentation[3].value.left_twips, Some(900));
    assert_eq!(indentation[3].value.first_line_twips, Some(120));
    assert_eq!(indentation[4].value.left_twips, Some(1417));
    assert_eq!(indentation[5].value.left_twips, Some(1417));
    assert_eq!(indentation[5].value.hanging_twips, Some(793));
    assert_eq!(indentation[6].value.left_twips, Some(1417));
    assert_eq!(indentation[6].value.hanging_twips, Some(180));
    assert_eq!(indentation[7].paragraph_ordinal, 7);
    assert_eq!(indentation[7].value.left_twips, Some(357));
    assert_eq!(indentation[7].value.right_twips, None);
    assert_eq!(indentation[7].value.first_line_twips, None);
    assert_eq!(indentation[7].value.hanging_twips, None);
    assert!(
        indentation.iter().all(|fact| fact.paragraph_ordinal != 8),
        "a cancelled numbered style without direct indentation must emit no indentation fact"
    );
}

#[test]
fn resolves_character_and_twip_indentation_without_conflating_zero_semantics() {
    let document = br#"
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="Twips"/><w:ind w:firstLineChars="0" w:startChars="0"/></w:pPr><w:r><w:t>Zero chars</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Chars"/></w:pPr><w:r><w:t>Character units</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Chars"/><w:ind w:firstLine="600" w:left="700"/></w:pPr><w:r><w:t>Inherited chars</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Chars"/><w:ind w:firstLine="600" w:left="700" w:firstLineChars="0" w:startChars="0"/></w:pPr><w:r><w:t>Cancelled chars</w:t></w:r></w:p>
        <w:p><w:pPr><w:ind w:firstLine="300" w:hanging="400"/></w:pPr><w:r><w:t>Hanging twips wins</w:t></w:r></w:p>
        <w:p><w:pPr><w:ind w:firstLineChars="100" w:hangingChars="200"/></w:pPr><w:r><w:t>Hanging chars wins</w:t></w:r></w:p>
      </w:body></w:document>
    "#;
    let styles = br#"
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Twips"><w:pPr><w:ind w:firstLine="300" w:left="400"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Chars"><w:basedOn w:val="Twips"/><w:pPr><w:ind w:firstLineChars="100" w:startChars="200"/></w:pPr></w:style>
      </w:styles>
    "#;
    let projection = project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .unwrap();
    let StructuralFactSet::Known(indentation) = projection.structural_facts.indentation else {
        panic!("effective indentation should be known");
    };

    assert_eq!(indentation[0].value.first_line_twips, Some(300));
    assert_eq!(indentation[0].value.left_twips, Some(400));
    assert_eq!(indentation[0].value.first_line_chars_hundredths, None);
    assert_eq!(indentation[0].value.start_chars_hundredths, None);
    for fact in [&indentation[1], &indentation[2]] {
        assert_eq!(fact.value.first_line_twips, None);
        assert_eq!(fact.value.left_twips, None);
        assert_eq!(fact.value.first_line_chars_hundredths, Some(100));
        assert_eq!(fact.value.start_chars_hundredths, Some(200));
    }
    assert_eq!(indentation[3].value.first_line_twips, Some(600));
    assert_eq!(indentation[3].value.left_twips, Some(700));
    assert_eq!(indentation[3].value.first_line_chars_hundredths, None);
    assert_eq!(indentation[3].value.start_chars_hundredths, None);
    assert_eq!(indentation[4].value.first_line_twips, None);
    assert_eq!(indentation[4].value.hanging_twips, Some(400));
    assert_eq!(indentation[5].value.first_line_chars_hundredths, None);
    assert_eq!(indentation[5].value.hanging_chars_hundredths, Some(200));
}

#[test]
fn invalid_or_missing_numbering_fails_closed_without_losing_other_structure() {
    let document: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let invalid_numbering: &[u8] = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="5"/></w:numbering>"#;
    for entries in [
        vec![
            ("word/document.xml", document),
            ("word/styles.xml", MINIMAL_STYLES),
        ],
        vec![
            ("word/document.xml", document),
            ("word/styles.xml", MINIMAL_STYLES),
            ("word/numbering.xml", invalid_numbering),
        ],
    ] {
        let projection = project_docx(
            &package(&entries, CompressionMethod::Deflated),
            DocxLimits::default(),
            allocate,
        )
        .expect("unsupported optional numbering should preserve the document");
        assert_eq!(
            projection.structural_facts.indentation,
            StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering)
        );
        assert_eq!(
            projection.structural_facts.outline_levels,
            StructuralFactSet::Known(vec![ParagraphOutlineLevelFact {
                paragraph_ordinal: 0,
                outline_level: 2,
            }])
        );
        assert_eq!(
            projection.structural_facts.numbering_hierarchy,
            StructuralFactSet::Known(Vec::new())
        );
    }
}

#[test]
fn enforces_numbering_package_and_item_limits() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
    let numbering = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#;
    let package = package(
        &[
            ("word/document.xml", document),
            ("word/styles.xml", MINIMAL_STYLES),
            ("word/numbering.xml", numbering),
        ],
        CompressionMethod::Deflated,
    );
    assert_eq!(
        project_docx(
            &package,
            DocxLimits {
                maximum_numbering_xml_bytes: numbering.len() - 1,
                ..DocxLimits::default()
            },
            allocate,
        ),
        Err(ProjectionError::NumberingXmlTooLarge)
    );
    assert_eq!(
        project_docx(
            &package,
            DocxLimits {
                maximum_numbering_items: 1,
                ..DocxLimits::default()
            },
            allocate,
        ),
        Err(ProjectionError::TooManyNumberingItems)
    );
}

#[test]
fn document_part_only_does_not_claim_style_dependent_facts() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p></w:body></w:document>"#;
    let projection = project_document_xml(xml, allocate).expect("document part should project");

    assert_eq!(
        projection.paragraphs[0].style_id.as_deref(),
        Some("Heading")
    );
    assert_eq!(
        projection.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::DocumentPartOnly)
    );
    assert_eq!(
        projection.structural_facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::DocumentPartOnly)
    );
    assert_eq!(
        projection.structural_facts.numbering_hierarchy,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::DocumentPartOnly)
    );
    assert_eq!(
        projection.structural_facts.outline_levels,
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
fn direct_style_ids_survive_unavailable_and_malformed_style_sheets() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let unavailable = project_docx(
        &package(
            &[("word/document.xml", document)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("a missing optional styles part should preserve the document");
    assert_eq!(
        unavailable.paragraphs[0].style_id.as_deref(),
        Some("Heading")
    );
    assert_eq!(
        unavailable.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::StylesPartUnavailable)
    );
    assert_eq!(
        unavailable.structural_facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::StylesPartUnavailable)
    );

    let malformed = project_docx(
        &package(
            &[
                ("word/document.xml", document),
                ("word/styles.xml", b"<not-styles/>"),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("an invalid optional styles part should preserve the document");
    assert_eq!(malformed.paragraphs[0].style_id.as_deref(), Some("Heading"));
    assert_eq!(
        malformed.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        malformed.structural_facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );

    let unknown_style = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("an unknown style reference should preserve the document");
    assert_eq!(
        unknown_style.paragraphs[0].style_id.as_deref(),
        Some("Heading")
    );
    assert_eq!(
        unknown_style.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        unknown_style.structural_facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
}

fn alignments(document: &[u8], styles: &[u8]) -> Vec<Option<ParagraphAlignmentFact>> {
    project_docx(
        &package(
            &[("word/document.xml", document), ("word/styles.xml", styles)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        allocate,
    )
    .expect("a supported style sheet should resolve paragraph alignment")
    .paragraphs
    .into_iter()
    .map(|paragraph| paragraph.alignment)
    .collect()
}

fn direct_alignments(document: &[u8]) -> Vec<Option<ParagraphAlignmentFact>> {
    project_document_xml(document, allocate)
        .expect("the document part should project")
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.alignment)
        .collect()
}

const fn direct_alignment(value: Align) -> ParagraphAlignmentFact {
    ParagraphAlignmentFact {
        value,
        source: ParagraphAlignmentSource::Direct,
    }
}

const fn style_alignment(value: Align) -> ParagraphAlignmentFact {
    ParagraphAlignmentFact {
        value,
        source: ParagraphAlignmentSource::Style,
    }
}

#[test]
fn paragraph_alignment_prefers_direct_over_style_over_document_defaults() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Body"/><w:jc w:val="right"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Plain"/></w:pPr><w:r><w:t>Defaulted</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let cascaded = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
      <w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Plain"/>
    </w:styles>"#;
    assert_eq!(
        alignments(document, cascaded),
        [
            Some(direct_alignment(Align::Right)),
            Some(style_alignment(Align::Center)),
            Some(style_alignment(Align::Justify)),
        ]
    );

    let without_document_defaults =
        br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
      <w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Plain"/>
    </w:styles>"#;
    assert_eq!(
        alignments(document, without_document_defaults),
        [
            Some(direct_alignment(Align::Right)),
            Some(style_alignment(Align::Center)),
            None,
        ]
    );
}

#[test]
fn paragraph_alignment_resolves_based_on_chains_through_cycles_and_missing_styles() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Level3"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Override"/></w:pPr><w:r><w:t>Overridden</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="CycleA"/></w:pPr><w:r><w:t>Cyclic</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Absent"/></w:pPr><w:r><w:t>Missing</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let styles = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
      <w:style w:type="paragraph" w:styleId="Level1"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Level2"><w:basedOn w:val="Level1"/></w:style>
      <w:style w:type="paragraph" w:styleId="Level3"><w:basedOn w:val="Level2"/></w:style>
      <w:style w:type="paragraph" w:styleId="Override"><w:basedOn w:val="Level1"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style>
    </w:styles>"#;
    // An unresolvable initial style projects no alignment: Word falls back to
    // Normal there, so `w:docDefaults` would report a value the document never
    // resolves to.
    assert_eq!(
        alignments(document, styles),
        [
            Some(style_alignment(Align::Center)),
            Some(style_alignment(Align::Right)),
            None,
            None,
        ]
    );
}

#[test]
fn paragraph_alignment_uses_the_default_paragraph_style_when_no_style_is_named() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>Plain</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let styles = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
    </w:styles>"#;
    assert_eq!(
        alignments(document, styles),
        [Some(style_alignment(Align::Center))]
    );
    assert_eq!(alignments(document, MINIMAL_STYLES), [None]);
}

#[test]
fn paragraph_alignment_maps_justification_tokens_and_shadows_unsupported_ones() {
    for (token, expected) in [
        ("center", Some(direct_alignment(Align::Center))),
        ("both", Some(direct_alignment(Align::Justify))),
        ("distribute", Some(direct_alignment(Align::Justify))),
        ("left", Some(direct_alignment(Align::Left))),
        ("right", Some(direct_alignment(Align::Right))),
        // `start` and `end` depend on the paragraph's reading order, which is
        // not resolved here.
        ("start", None),
        ("end", None),
        ("thaiDistribute", None),
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="{token}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
        );
        assert_eq!(
            direct_alignments(xml.as_bytes()),
            [expected],
            "w:jc {token}"
        );
    }

    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Body"/><w:jc w:val="thaiDistribute"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Exotic"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Body"/><w:jc w:val="start"/></w:pPr><w:r><w:t>Logical direct</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Logical"/></w:pPr><w:r><w:t>Logical style</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let styles = br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
      <w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Exotic"><w:pPr><w:jc w:val="thaiDistribute"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Logical"><w:pPr><w:jc w:val="end"/></w:pPr></w:style>
    </w:styles>"#;
    assert_eq!(alignments(document, styles), [None, None, None, None]);
}

#[test]
fn table_cell_paragraph_alignment_ignores_justification_outside_paragraph_properties() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:jc w:val="center"/></w:tblPr>
      <w:tr>
        <w:trPr><w:jc w:val="center"/></w:trPr>
        <w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:jc w:val="center"/><w:r><w:jc w:val="center"/><w:t>Plain</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl></w:body></w:document>"#;
    let styles =
        br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
      <w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:jc w:val="left"/></w:pPr></w:style>
    </w:styles>"#;
    assert_eq!(
        alignments(document, styles),
        [
            Some(direct_alignment(Align::Right)),
            Some(style_alignment(Align::Left)),
            None,
        ]
    );
}

#[test]
fn paragraph_alignment_ignores_a_paragraph_property_change_snapshot() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Current</w:t></w:r></w:p>
      <w:p><w:pPr><w:pPrChange w:id="2" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Cleared</w:t></w:r></w:p>
    </w:body></w:document>"#;
    assert_eq!(
        alignments(document, MINIMAL_STYLES),
        [Some(direct_alignment(Align::Center)), None]
    );
}

#[test]
fn degraded_style_sheets_project_direct_paragraph_alignment_only() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Body"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let expected = [Some(direct_alignment(Align::Center)), None];

    let document_part_only =
        project_document_xml(document, allocate).expect("the document part should project");
    assert_eq!(
        document_part_only
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.alignment)
            .collect::<Vec<_>>(),
        expected
    );
    assert_eq!(
        document_part_only.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::DocumentPartOnly)
    );

    for (styles, reason) in [
        (None, FormattingUnknownReason::StylesPartUnavailable),
        (
            Some(&b"<not-styles/>"[..]),
            FormattingUnknownReason::UnsupportedStyles,
        ),
    ] {
        let mut entries = vec![("word/document.xml", &document[..])];
        if let Some(styles) = styles {
            entries.push(("word/styles.xml", styles));
        }
        let projection = project_docx(
            &package(&entries, CompressionMethod::Deflated),
            DocxLimits::default(),
            allocate,
        )
        .expect("a degraded styles part should preserve the document");
        assert_eq!(
            projection
                .paragraphs
                .iter()
                .map(|paragraph| paragraph.alignment)
                .collect::<Vec<_>>(),
            expected,
            "{reason:?}"
        );
        assert_eq!(
            projection.formatting_status,
            FormattingProjectionStatus::Incomplete(reason)
        );
    }
}

#[test]
fn paragraph_alignment_resolves_under_strict_and_transitional_namespaces() {
    for namespace in [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let document = format!(
            r#"<a:document xmlns:a="{namespace}"><a:body><a:p><a:pPr><a:pStyle a:val="Body"/></a:pPr><a:r><a:t>Styled</a:t></a:r></a:p><a:p><a:pPr><a:jc a:val="right"/></a:pPr><a:r><a:t>Direct</a:t></a:r></a:p></a:body></a:document>"#,
        );
        let styles = format!(
            r#"<z:styles xmlns:z="{namespace}"><z:style z:type="paragraph" z:styleId="Normal" z:default="1"/><z:style z:type="paragraph" z:styleId="Body"><z:pPr><z:jc z:val="distribute"/></z:pPr></z:style></z:styles>"#,
        );
        assert_eq!(
            alignments(document.as_bytes(), styles.as_bytes()),
            [
                Some(style_alignment(Align::Justify)),
                Some(direct_alignment(Align::Right)),
            ],
            "{namespace}"
        );
    }
}

#[test]
fn valid_style_sheets_make_absent_outline_levels_authoritatively_empty() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("a supported style sheet should resolve outline facts");

    assert_eq!(
        projection.structural_facts.outline_levels,
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
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedNumbering)
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
        cyclic_style_projection.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        cyclic_style_projection.structural_facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        cyclic_style_projection.structural_facts.numbering_hierarchy,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        cyclic_style_projection.structural_facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::UnsupportedStyles)
    );
    assert_eq!(
        cyclic_style_projection.paragraphs[0].style_id.as_deref(),
        Some("CycleA"),
        "unsupported inheritance must not discard direct paragraph facts"
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
fn projects_bookmarks_at_table_row_boundaries() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:tbl><w:tr>
        <w:bookmarkStart w:id="4" w:name="RowRange"/>
        <w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>
        <w:bookmarkEnd w:id="4"/>
      </w:tr></w:tbl>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("row-level bookmark boundaries should project");

    let StructuralFactSet::Known(bookmarks) = projection.structural_facts.bookmarks else {
        panic!("row-level bookmark boundaries should remain complete");
    };
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].paragraph_ordinal, 0);
    assert_eq!(bookmarks[0].span.start_utf8, 0);
    assert_eq!(bookmarks[0].span.end_utf8, 5);
    assert_eq!(bookmarks[0].span.coverage, SpanCoverage::Complete);
}

#[test]
fn projects_bookmarks_between_table_cells_and_across_rows() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:tbl>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:bookmarkStart w:id="7" w:name="CrossRow"/>
          <w:tc><w:p><w:r><w:t>Bravo</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Charlie</w:t></w:r></w:p></w:tc>
          <w:bookmarkEnd w:id="7"/>
          <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("cross-row bookmark boundaries should project");

    let StructuralFactSet::Known(bookmarks) = projection.structural_facts.bookmarks else {
        panic!("cross-row bookmark boundaries should remain complete");
    };
    assert_eq!(bookmarks.len(), 2);
    assert_eq!(bookmarks[0].paragraph_ordinal, 1);
    assert_eq!(bookmarks[0].span.start_utf8, 0);
    assert_eq!(bookmarks[0].span.end_utf8, 5);
    assert_eq!(bookmarks[0].span.coverage, SpanCoverage::ContinuesAfter);
    assert_eq!(bookmarks[1].paragraph_ordinal, 2);
    assert_eq!(bookmarks[1].span.start_utf8, 0);
    assert_eq!(bookmarks[1].span.end_utf8, 7);
    assert_eq!(bookmarks[1].span.coverage, SpanCoverage::ContinuesBefore);
}

#[test]
fn projects_body_and_zero_width_bookmark_boundaries() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:bookmarkStart w:id="4" w:name="BodyRange"/>
      <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="4"/>
      <w:bookmarkStart w:id="7" w:name="Cursor"/><w:bookmarkEnd w:id="7"/>
      <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("body-level and zero-width bookmark boundaries should project");

    let StructuralFactSet::Known(bookmarks) = projection.structural_facts.bookmarks else {
        panic!("body-level and zero-width bookmark boundaries should remain complete");
    };
    assert_eq!(bookmarks.len(), 2);
    assert_eq!(bookmarks[0].name, "BodyRange");
    assert_eq!(bookmarks[0].paragraph_ordinal, 0);
    assert_eq!(bookmarks[0].span.start_utf8, 0);
    assert_eq!(bookmarks[0].span.end_utf8, 5);
    assert_eq!(bookmarks[1].name, "Cursor");
    assert_eq!(bookmarks[1].paragraph_ordinal, 1);
    assert_eq!(bookmarks[1].span.start_utf8, 0);
    assert_eq!(bookmarks[1].span.end_utf8, 0);
    assert_eq!(bookmarks[1].span.coverage, SpanCoverage::Complete);
}

#[test]
fn malformed_bookmark_boundaries_fail_the_whole_fact_family_closed() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:bookmarkStart w:id="4" w:name="OtherwiseValid"/>
      <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="4"/>
      <w:bookmarkEnd w:id="999"/>
    </w:body></w:document>"#;
    let projection = project_docx(
        &package_with_minimal_styles(document),
        DocxLimits::default(),
        allocate,
    )
    .expect("malformed bookmark boundaries should preserve paragraph projection");

    assert_eq!(
        projection.structural_facts.bookmarks,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::IncompleteBookmarkRanges),
        "one malformed range must invalidate every bookmark fact"
    );
}

#[test]
fn package_projection_is_deterministic_and_styles_extraction_is_bounded() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Stable</w:t></w:r></w:p></w:body></w:document>"#;
    let package_bytes = package_with_minimal_styles(document);
    let first = project_docx(&package_bytes, DocxLimits::default(), allocate).unwrap();
    let second = project_docx(&package_bytes, DocxLimits::default(), allocate).unwrap();
    assert_eq!(first, second);
    assert_eq!(
        first.formatting_status,
        FormattingProjectionStatus::Complete
    );
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

    let two_styles =
        br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="TableOne"/>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
    </w:styles>"#;
    let style_count_package = package(
        &[
            ("word/document.xml", document),
            ("word/styles.xml", two_styles),
        ],
        CompressionMethod::Deflated,
    );
    let accepted = project_docx(
        &style_count_package,
        DocxLimits {
            maximum_styles: 2,
            ..DocxLimits::default()
        },
        allocate,
    )
    .expect("the exact style-count boundary should be accepted");
    assert_eq!(
        accepted.formatting_status,
        FormattingProjectionStatus::Complete
    );
    let rejected = project_docx(
        &style_count_package,
        DocxLimits {
            maximum_styles: 1,
            ..DocxLimits::default()
        },
        allocate,
    )
    .expect("an oversized optional styles part should preserve direct document facts");
    assert_eq!(
        rejected.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
}

#[test]
#[allow(clippy::too_many_lines)] // One end-to-end fixture asserts the fused projection contract.
fn fuses_document_projection_with_attributed_revisions_and_comment_threads() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:commentRangeStart w:id="19"/><w:ins w:id="7" w:author="Ada" w:date="2026-07-01T10:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins><w:commentRangeEnd w:id="19"/><w:r><w:commentReference w:id="19"/></w:r><w:del w:id="8" w:author="Lin"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>
    </w:body></w:document>"#;
    let comments_xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:comment w:id="19" w:author="Ada" w:initials="AA"><w:p w14:paraId="AAAAAAAA"><w:r><w:t>First</w:t></w:r></w:p><w:p w14:paraId="1F21D71D"><w:r><w:t>second</w:t></w:r></w:p></w:comment>
      <w:comment w:id="20" w:author="Lin"><w:p w14:paraId="0D537B10"/></w:comment>
      <w:comment w:id="22" w:author="Mae"><w:p w14:paraId="0ED2E4B2"/></w:comment>
    </w:comments>"#;
    let extended =
        br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
      <w15:commentEx w15:paraId="1F21D71D"/>
      <w15:commentEx w15:paraId="0D537B10" w15:paraIdParent="1F21D71D" w15:done="0"/>
      <w15:commentEx w15:paraId="0ED2E4B2" w15:done="1"/>
    </w15:commentsEx>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[
                ("word/document.xml", document),
                ("word/comments.xml", comments_xml),
                ("word/commentsExtended.xml", extended),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();

    assert_eq!(projection.document.paragraphs[0].text, "new");
    let ReviewFactSet::Known(revisions) = projection.review_facts.revisions else {
        panic!("valid revisions should be complete");
    };
    assert_eq!(revisions.len(), 2);
    assert_eq!(revisions[0].kind, RevisionFactKind::Insertion);
    assert_eq!(revisions[0].author, "Ada");
    assert_eq!(revisions[0].revision_id.as_deref(), Some("7"));
    assert_eq!(
        revisions[0].content,
        ReviewDetail::Known(RevisionContent {
            span: ReviewSpan {
                start: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 0,
                    utf16: 0,
                },
                end: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 3,
                    utf16: 3,
                },
            },
            text: "new".to_owned(),
            formatting_only: false,
        })
    );
    assert_eq!(revisions[1].kind, RevisionFactKind::Deletion);
    assert_eq!(
        revisions[1].content,
        ReviewDetail::Known(RevisionContent {
            span: ReviewSpan {
                start: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 3,
                    utf16: 3,
                },
                end: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 3,
                    utf16: 3,
                },
            },
            text: "old".to_owned(),
            formatting_only: false,
        })
    );

    let ReviewFactSet::Known(comments) = projection.review_facts.comments else {
        panic!("valid comments should be complete");
    };
    assert_eq!(comments.len(), 3);
    assert_eq!(
        comments[0].content,
        ReviewDetail::Known(CommentContent {
            anchor: ReviewSpan {
                start: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 0,
                    utf16: 0,
                },
                end: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 3,
                    utf16: 3,
                },
            },
            comment_text: "First\nsecond".to_owned(),
            referenced_text: "new".to_owned(),
        })
    );
    assert_eq!(comments[1].comment_id, "20");
    assert_eq!(comments[1].parent_comment_id.as_deref(), Some("19"));
    assert!(!comments[1].resolved);
    assert!(comments[2].resolved);
}

#[test]
fn bounds_aggregate_review_detail_bytes_before_repeating_anchor_text() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/><w:ins w:id="3"><w:r><w:t>0123456789</w:t></w:r></w:ins><w:commentRangeEnd w:id="1"/><w:commentRangeEnd w:id="2"/></w:p>
    </w:body></w:document>"#;
    let comments =
        br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:comment w:id="1"><w:p><w:r><w:t>A</w:t></w:r></w:p></w:comment>
      <w:comment w:id="2"><w:p><w:r><w:t>B</w:t></w:r></w:p></w:comment>
    </w:comments>"#;
    let package_bytes = package(
        &[
            ("word/document.xml", document),
            ("word/comments.xml", comments),
        ],
        CompressionMethod::Deflated,
    );
    let project = |maximum_review_detail_bytes| {
        project_docx_with_review_facts(
            &package_bytes,
            DocxLimits::default(),
            ReviewFactLimits {
                maximum_review_detail_bytes,
                ..ReviewFactLimits::default()
            },
            ProjectionOptions::default(),
            allocate,
        )
        .unwrap()
    };

    let ReviewFactSet::Known(at_boundary) = project(32).review_facts.comments else {
        panic!("the exact aggregate output boundary must remain known");
    };
    assert_eq!(at_boundary.len(), 2);
    assert_eq!(
        project(31).review_facts.comments,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::ResourceLimit)
    );
}

#[test]
fn whitespace_only_revision_content_is_not_formatting_only() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:ins w:id="1"><w:r><w:t xml:space="preserve"> </w:t><w:tab/><w:br/><w:t>&#xA0;&#x2003;</w:t></w:r></w:ins>
    </w:p></w:body></w:document>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[("word/document.xml", document)],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    let ReviewFactSet::Known(revisions) = projection.review_facts.revisions else {
        panic!("bounded revision facts should be known");
    };
    let ReviewDetail::Known(content) = &revisions[0].content else {
        panic!("inline revision content should be known");
    };
    assert_eq!(content.text, " \t\u{000b}\u{00a0}\u{2003}");
    assert!(!content.formatting_only);
}

#[test]
fn nested_revision_text_populates_every_enclosing_fact_under_one_budget() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:ins w:id="1" w:author="Outer"><w:r><w:t>A</w:t></w:r><w:del w:id="2" w:author="Inner"><w:r><w:delText>B</w:delText></w:r></w:del><w:r><w:t>C</w:t></w:r></w:ins>
    </w:p></w:body></w:document>"#;
    let package_bytes = package(
        &[("word/document.xml", document)],
        CompressionMethod::Deflated,
    );
    let project = |maximum_review_detail_bytes| {
        project_docx_with_review_facts(
            &package_bytes,
            DocxLimits::default(),
            ReviewFactLimits {
                maximum_review_detail_bytes,
                ..ReviewFactLimits::default()
            },
            ProjectionOptions::default(),
            allocate,
        )
        .unwrap()
    };

    let ReviewFactSet::Known(revisions) = project(4).review_facts.revisions else {
        panic!("the exact nested revision output boundary should remain known");
    };
    assert_eq!(revisions.len(), 2);
    let ReviewDetail::Known(outer) = &revisions[0].content else {
        panic!("the outer revision should have textual content");
    };
    let ReviewDetail::Known(inner) = &revisions[1].content else {
        panic!("the inner revision should have textual content");
    };
    assert_eq!(outer.text, "ABC");
    assert_eq!(inner.text, "B");
    assert!(!outer.formatting_only);
    assert!(!inner.formatting_only);
    assert_eq!(
        project(3).review_facts.revisions,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::ResourceLimit)
    );
}

#[test]
fn comment_controls_follow_the_requested_text_materialization() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>X</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p></w:body></w:document>"#;
    let comments_xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="8640"/></w:tabs><w:rPr><w:tab/></w:rPr></w:pPr><w:r><w:rPr><w:tab/></w:rPr><w:t>A</w:t><w:tab/><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:cr/><w:softHyphen/><w:noBreakHyphen/><w:t>Z</w:t></w:r></w:p></w:comment></w:comments>"#;
    let package_bytes = package(
        &[
            ("word/document.xml", document),
            ("word/comments.xml", comments_xml),
        ],
        CompressionMethod::Deflated,
    );
    for (materialization, expected) in [
        (
            TextMaterialization::WordHost,
            "A\t\u{000b}\u{000c}\u{000e}\r\u{001f}\u{001e}Z",
        ),
        (
            TextMaterialization::ReadablePlainText,
            "A\t\n\n\n\n\u{00ad}\u{2011}Z",
        ),
    ] {
        let projection = project_docx_with_review_facts(
            &package_bytes,
            DocxLimits::default(),
            ReviewFactLimits::default(),
            ProjectionOptions {
                text_materialization: materialization,
                ..ProjectionOptions::default()
            },
            allocate,
        )
        .unwrap();
        let ReviewFactSet::Known(comments) = projection.review_facts.comments else {
            panic!("valid comments should be known");
        };
        let ReviewDetail::Known(content) = &comments[0].content else {
            panic!("anchored comment content should be known");
        };
        assert_eq!(content.comment_text, expected);
    }
}

#[test]
fn comment_markers_inside_suppressed_textboxes_cannot_create_known_anchors() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:r><w:t>A</w:t></w:r><w:txbxContent><w:p><w:commentRangeStart w:id="1"/><w:ins w:id="2" w:author="Reviewer"><w:r><w:t>hidden</w:t><w:commentReference w:id="1"/></w:r></w:ins><w:commentRangeEnd w:id="1"/></w:p></w:txbxContent><w:r><w:t>B</w:t></w:r>
    </w:p></w:body></w:document>"#;
    let comments_xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"><w:p><w:r><w:t>Review</w:t></w:r></w:p></w:comment></w:comments>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[
                ("word/document.xml", document),
                ("word/comments.xml", comments_xml),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    assert_eq!(projection.document.paragraphs[0].text, "AB");
    let ReviewFactSet::Known(comments) = projection.review_facts.comments else {
        panic!("valid comments should be known");
    };
    assert_eq!(
        comments[0].content,
        ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation)
    );
    let ReviewFactSet::Known(revisions) = projection.review_facts.revisions else {
        panic!("valid revision attribution should be known");
    };
    assert_eq!(revisions.len(), 1);
    assert_eq!(revisions[0].author, "Reviewer");
    assert_eq!(
        revisions[0].content,
        ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation)
    );
}

#[test]
fn accepts_sparse_extensions_wrapper_ids_case_variants_and_full_on_off_values() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
    let comments = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:comment w:id="1" w14:paraId="a1b2c3d4"><w:p w14:paraId="FFFFFFFF"/></w:comment>
      <w:comment w:id="2"><w:p/></w:comment>
      <w:comment w:id="3"><w:p w14:paraId="e5f607a8"/></w:comment>
    </w:comments>"#;
    for (value, expected) in [
        ("1", true),
        ("true", true),
        ("ON", true),
        ("0", false),
        ("False", false),
        ("off", false),
    ] {
        let extended = format!(
            r#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
              <w15:commentEx w15:paraId="A1B2C3D4"/>
              <w15:commentEx w15:paraId="E5F607A8" w15:paraIdParent="A1b2C3d4" w15:done="{value}"/>
            </w15:commentsEx>"#
        );
        let projection = project_docx_with_review_facts(
            &package(
                &[
                    ("word/document.xml", document),
                    ("word/comments.xml", comments),
                    ("word/commentsExtended.xml", extended.as_bytes()),
                ],
                CompressionMethod::Deflated,
            ),
            DocxLimits::default(),
            ReviewFactLimits::default(),
            ProjectionOptions::default(),
            allocate,
        )
        .unwrap();
        let ReviewFactSet::Known(facts) = projection.review_facts.comments else {
            panic!("valid sparse comment extension maps must remain known");
        };
        assert_eq!(facts.len(), 3);
        assert_eq!(facts[2].parent_comment_id.as_deref(), Some("1"));
        assert_eq!(facts[2].resolved, expected);
        assert!(!facts[1].resolved);
    }
}

#[test]
fn records_every_supported_attributed_revision_kind_in_the_effective_body() {
    let tags = [
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "cellIns",
        "cellDel",
        "cellMerge",
        "pPrChange",
        "rPrChange",
        "sectPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "tblGridChange",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlMoveFromRangeStart",
        "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart",
        "customXmlMoveToRangeEnd",
    ];
    let mut markup = String::new();
    for (index, tag) in tags.iter().enumerate() {
        write!(markup, r#"<w:{tag} w:id="{index}" w:author="Reviewer"/>"#).unwrap();
    }
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>{markup}</w:p></w:body></w:document>"#
    );
    let projection = project_docx_with_review_facts(
        &package(
            &[("word/document.xml", document.as_bytes())],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    let ReviewFactSet::Known(revisions) = projection.review_facts.revisions else {
        panic!("all recognized revision markup must produce complete attribution");
    };
    assert_eq!(
        revisions
            .iter()
            .map(|revision| revision.kind)
            .collect::<Vec<_>>(),
        [
            RevisionFactKind::Insertion,
            RevisionFactKind::Deletion,
            RevisionFactKind::MoveFrom,
            RevisionFactKind::MoveTo,
            RevisionFactKind::CellInsertion,
            RevisionFactKind::CellDeletion,
            RevisionFactKind::CellMerge,
            RevisionFactKind::ParagraphPropertiesChange,
            RevisionFactKind::RunPropertiesChange,
            RevisionFactKind::SectionPropertiesChange,
            RevisionFactKind::TablePropertiesChange,
            RevisionFactKind::TableRowPropertiesChange,
            RevisionFactKind::TableCellPropertiesChange,
            RevisionFactKind::TableGridChange,
            RevisionFactKind::CustomXmlDeletionRangeStart,
            RevisionFactKind::CustomXmlDeletionRangeEnd,
            RevisionFactKind::CustomXmlInsertionRangeStart,
            RevisionFactKind::CustomXmlInsertionRangeEnd,
            RevisionFactKind::CustomXmlMoveFromRangeStart,
            RevisionFactKind::CustomXmlMoveFromRangeEnd,
            RevisionFactKind::CustomXmlMoveToRangeStart,
            RevisionFactKind::CustomXmlMoveToRangeEnd,
        ]
    );
    assert!(
        revisions
            .iter()
            .all(|revision| revision.author == "Reviewer")
    );
}

#[test]
fn resolves_review_parts_only_through_document_relationships() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
    let comments = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Ada"/></w:comments>"#;
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="review" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/comments" Target="review/comments-custom.xml"/></Relationships>"#;
    let package_relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/main.xml"/></Relationships>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[
                ("_rels/.rels", package_relationships),
                ("custom/main.xml", document),
                ("custom/_rels/main.xml.rels", relationships),
                ("custom/review/comments-custom.xml", comments),
                ("word/comments.xml", b"not XML"),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    let ReviewFactSet::Known(projected_comments) = projection.review_facts.comments else {
        panic!("relationship-selected comments should be known");
    };
    assert_eq!(projected_comments.len(), 1);
    assert_eq!(projected_comments[0].comment_id, "7");
}

#[test]
fn projects_whole_and_multiple_paragraph_comment_ranges_without_cross_id_poisoning() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:commentRangeStart w:id="1"/><w:p><w:r><w:t>whole&#x1F600;</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p><w:commentRangeEnd w:id="1"/>
      <w:p><w:commentRangeStart w:id="2"/><w:commentRangeStart w:id="2"/><w:r><w:t>invalid</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>
    </w:body></w:document>"#;
    let comments_xml =
        br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:comment w:id="1" w:author="Ada"><w:p><w:r><w:t>Valid</w:t></w:r></w:p></w:comment>
      <w:comment w:id="2" w:author="Lin"><w:p><w:r><w:t>Invalid anchor</w:t></w:r></w:p></w:comment>
    </w:comments>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[
                ("word/document.xml", document),
                ("word/comments.xml", comments_xml),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();

    let ReviewFactSet::Known(comments) = projection.review_facts.comments else {
        panic!("valid comments should be complete");
    };
    assert_eq!(
        comments[0].content,
        ReviewDetail::Known(CommentContent {
            anchor: ReviewSpan {
                start: ReviewPoint {
                    paragraph_ordinal: 0,
                    utf8: 0,
                    utf16: 0,
                },
                end: ReviewPoint {
                    paragraph_ordinal: 1,
                    utf8: 6,
                    utf16: 6,
                },
            },
            comment_text: "Valid".to_owned(),
            referenced_text: "whole😀\nsecond".to_owned(),
        })
    );
    assert_eq!(
        comments[1].content,
        ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation)
    );
}

#[test]
fn missing_comments_are_authoritatively_empty() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
    let projection = project_docx_with_review_facts(
        &package(
            &[
                ("word/document.xml", document),
                ("word/commentsExtended.xml", b"not XML"),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    assert_eq!(
        projection.review_facts.comments,
        ReviewFactSet::Known(Vec::new())
    );
}

#[test]
fn malformed_or_incomplete_comment_graphs_fail_closed_without_losing_the_document() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p></w:body></w:document>"#;
    let comment = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="1"><w:p w14:paraId="AAAAAAAA"/></w:comment></w:comments>"#;
    let cases: [(&[u8], ReviewFactUnknownReason); 5] = [
        (b"<broken", ReviewFactUnknownReason::InvalidCommentsExtended),
        (
            br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="BBBBBBBB"/></w15:commentsEx>"#,
            ReviewFactUnknownReason::InvalidCommentsExtended,
        ),
        (
            br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="AAAAAAAA" w15:paraIdParent="BBBBBBBB"/></w15:commentsEx>"#,
            ReviewFactUnknownReason::InvalidCommentsExtended,
        ),
        (
            br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="AAAAAAAA" w15:paraIdParent="AAAAAAAA"/></w15:commentsEx>"#,
            ReviewFactUnknownReason::InvalidCommentsExtended,
        ),
        (
            br#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="AAAAAAAA" w15:done="perhaps"/></w15:commentsEx>"#,
            ReviewFactUnknownReason::InvalidCommentsExtended,
        ),
    ];
    for (extended, expected) in cases {
        let projection = project_docx_with_review_facts(
            &package(
                &[
                    ("word/document.xml", document),
                    ("word/comments.xml", comment),
                    ("word/commentsExtended.xml", extended),
                ],
                CompressionMethod::Deflated,
            ),
            DocxLimits::default(),
            ReviewFactLimits::default(),
            ProjectionOptions::default(),
            allocate,
        )
        .unwrap();
        assert_eq!(projection.document.paragraphs[0].text, "kept");
        assert_eq!(
            projection.review_facts.comments,
            ReviewFactSet::Unknown(expected)
        );
    }

    for malformed_comments in [
        &b"<broken"[..],
        &b"<wrong-root/>"[..],
        &br#"<!DOCTYPE w:comments [<!ENTITY x "value">]><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#[..],
    ] {
        let projection = project_docx_with_review_facts(
            &package(
                &[
                    ("word/document.xml", document),
                    ("word/comments.xml", malformed_comments),
                ],
                CompressionMethod::Deflated,
            ),
            DocxLimits::default(),
            ReviewFactLimits::default(),
            ProjectionOptions::default(),
            allocate,
        )
        .unwrap();
        assert_eq!(
            projection.review_facts.comments,
            ReviewFactSet::Unknown(ReviewFactUnknownReason::InvalidComments)
        );
    }
}

#[test]
fn duplicate_comment_parts_and_fact_limits_are_explicit_unknowns() {
    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins/><w:del/></w:p></w:body></w:document>"#;
    let comments =
        br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let mut package_bytes = package(
        &[
            ("word/document.xml", document),
            ("word/comments.xml", comments),
            ("word/commentz.xml", comments),
        ],
        CompressionMethod::Deflated,
    );
    replace_all_equal_length(
        &mut package_bytes,
        b"word/commentz.xml",
        b"word/comments.xml",
    );
    let projection = project_docx_with_review_facts(
        &package_bytes,
        DocxLimits::default(),
        ReviewFactLimits {
            maximum_facts_per_family: 1,
            ..ReviewFactLimits::default()
        },
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    assert_eq!(
        projection.review_facts.revisions,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::ResourceLimit)
    );
    assert_eq!(
        projection.review_facts.comments,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::InvalidComments)
    );

    let size_limited = project_docx_with_review_facts(
        &package(
            &[
                ("word/document.xml", document),
                ("word/comments.xml", comments),
            ],
            CompressionMethod::Deflated,
        ),
        DocxLimits::default(),
        ReviewFactLimits {
            maximum_comments_xml_bytes: 1,
            ..ReviewFactLimits::default()
        },
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    assert_eq!(
        size_limited.review_facts.comments,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::ResourceLimit)
    );

    let mut dishonest_size = package(
        &[
            ("word/document.xml", document),
            ("word/comments.xml", comments),
        ],
        CompressionMethod::Deflated,
    );
    set_zip_entry_uncompressed_size(&mut dishonest_size, b"word/comments.xml", 1);
    let overflow = project_docx_with_review_facts(
        &dishonest_size,
        DocxLimits::default(),
        ReviewFactLimits {
            maximum_comments_xml_bytes: comments.len().saturating_sub(1),
            ..ReviewFactLimits::default()
        },
        ProjectionOptions::default(),
        allocate,
    )
    .unwrap();
    assert_eq!(
        overflow.review_facts.comments,
        ReviewFactSet::Unknown(ReviewFactUnknownReason::ResourceLimit)
    );
}

proptest::proptest! {
    #[test]
    fn arbitrary_vertical_alignment_values_only_project_superscript(
        value in proptest::string::string_regex("[A-Za-z]{0,20}").unwrap()
    ) {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:vertAlign w:val="{value}"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let projection = project_document_xml(xml.as_bytes(), allocate)
            .expect("generated OOXML should project");
        let expected = if value == "superscript" {
            vec![TextFormattingSpan {
                start_utf16: 0,
                end_utf16: 1,
                style: TextStyle::Superscript,
            }]
        } else {
            Vec::new()
        };
        proptest::prop_assert_eq!(&projection.paragraphs[0].formatting, &expected);
    }

    #[test]
    fn arbitrary_justification_values_only_project_supported_alignments(
        value in proptest::string::string_regex(r"[\p{L}\p{N} _:.-]{0,64}").unwrap()
    ) {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="{value}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let projection = project_document_xml(xml.as_bytes(), allocate)
            .expect("generated OOXML should project");
        let expected = match value.as_str() {
            "center" => Some(direct_alignment(Align::Center)),
            "both" | "distribute" => Some(direct_alignment(Align::Justify)),
            "left" => Some(direct_alignment(Align::Left)),
            "right" => Some(direct_alignment(Align::Right)),
            _ => None,
        };
        proptest::prop_assert_eq!(projection.paragraphs[0].alignment, expected);
    }

    #[test]
    fn arbitrary_comments_never_abort_a_valid_document(comments in proptest::collection::vec(proptest::prelude::any::<u8>(), 0..2048)) {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>stable</w:t></w:r></w:p></w:body></w:document>"#;
        let package_bytes = package(
            &[("word/document.xml", document), ("word/comments.xml", &comments)],
            CompressionMethod::Deflated,
        );
        let projection = project_docx_with_review_facts(
            &package_bytes,
            DocxLimits::default(),
            ReviewFactLimits::default(),
            ProjectionOptions::default(),
            allocate,
        ).expect("optional comment bytes must not abort document projection");
        proptest::prop_assert_eq!(projection.document.paragraphs[0].text.as_str(), "stable");
    }

    #[test]
    fn arbitrary_numbering_never_aborts_a_valid_document(numbering in proptest::collection::vec(proptest::prelude::any::<u8>(), 0..2048)) {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>stable</w:t></w:r></w:p></w:body></w:document>"#;
        let package_bytes = package(
            &[
                ("word/document.xml", document),
                ("word/styles.xml", MINIMAL_STYLES),
                ("word/numbering.xml", &numbering),
            ],
            CompressionMethod::Deflated,
        );
        let projection = project_docx(&package_bytes, DocxLimits::default(), allocate)
            .expect("optional numbering bytes must not abort document projection");
        proptest::prop_assert_eq!(projection.paragraphs[0].text.as_str(), "stable");
        proptest::prop_assert!(matches!(
            projection.structural_facts.numbering_hierarchy,
            StructuralFactSet::Known(_)
        ));
    }
}

fn package_with_minimal_styles(document: &[u8]) -> Vec<u8> {
    package(
        &[
            ("word/document.xml", document),
            ("word/styles.xml", MINIMAL_STYLES),
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
    let has_document_relationships = entries
        .iter()
        .any(|(name, _)| *name == "word/_rels/document.xml.rels");
    if !has_document_relationships {
        let mut relationships = Vec::new();
        if entries.iter().any(|(name, _)| *name == "word/styles.xml") {
            relationships.push(
                r#"<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
            );
        }
        if entries
            .iter()
            .any(|(name, _)| *name == "word/numbering.xml")
        {
            relationships.push(
                r#"<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>"#,
            );
        }
        if entries.iter().any(|(name, _)| *name == "word/comments.xml") {
            relationships.push(
                r#"<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>"#,
            );
        }
        if entries
            .iter()
            .any(|(name, _)| *name == "word/commentsExtended.xml")
        {
            relationships.push(
                r#"<Relationship Id="commentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>"#,
            );
        }
        if !relationships.is_empty() {
            archive
                .start_file("word/_rels/document.xml.rels", options)
                .expect("test relationships entry should start");
            archive
                .write_all(
                    format!(
                        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
                        relationships.join("")
                    )
                    .as_bytes(),
                )
                .expect("test relationships entry should write");
        }
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

fn set_zip_entry_uncompressed_size(bytes: &mut [u8], expected_name: &[u8], size: u32) {
    let mut offset = 0;
    while offset + 46 <= bytes.len() {
        let (name_length_offset, name_offset, size_offset) =
            if bytes[offset..].starts_with(b"PK\x03\x04") {
                (offset + 26, offset + 30, offset + 22)
            } else if bytes[offset..].starts_with(b"PK\x01\x02") {
                (offset + 28, offset + 46, offset + 24)
            } else {
                offset += 1;
                continue;
            };
        let name_length = usize::from(u16::from_le_bytes([
            bytes[name_length_offset],
            bytes[name_length_offset + 1],
        ]));
        let name_end = name_offset.saturating_add(name_length);
        if bytes.get(name_offset..name_end) == Some(expected_name) {
            bytes[size_offset..size_offset + 4].copy_from_slice(&size.to_le_bytes());
        }
        offset = name_end;
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
