#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use stella_docx_kernel::{
    DocumentProjection, DocxLimits, FormattingProjectionStatus, FormattingUnknownReason,
    InternalParagraphId, ParagraphIdentityFacts, ProjectionError, StructuralFactSet,
    StructuralFactUnknownReason, project_docx, project_paragraph_fragment,
};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const PACKAGE_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/xmlPackage";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIP_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("paragraph-{}", facts.ordinal))
}

fn document(body: &str) -> String {
    format!(r#"<w:document xmlns:w="{WORD_NAMESPACE}"><w:body>{body}</w:body></w:document>"#)
}

fn root_relationships() -> String {
    format!(
        r#"<Relationships xmlns="{RELATIONSHIP_NAMESPACE}"><Relationship Type="{OFFICE_NAMESPACE}/officeDocument" Target="custom/main.xml"/></Relationships>"#
    )
}

fn dependency_relationships() -> String {
    format!(
        r#"<Relationships xmlns="{RELATIONSHIP_NAMESPACE}"><Relationship Type="{OFFICE_NAMESPACE}/styles" Target="../shared/styles.xml"/><Relationship Type="{OFFICE_NAMESPACE}/numbering" Target="../shared/numbering.xml"/></Relationships>"#
    )
}

fn styles() -> String {
    format!(
        r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:b/></w:rPr></w:style></w:styles>"#
    )
}

fn numbering() -> String {
    format!(
        r#"<w:numbering xmlns:w="{WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"/><w:lvl w:ilvl="1"/></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num></w:numbering>"#
    )
}

fn parts(main: &str, dependencies: bool) -> Vec<(String, String)> {
    let mut parts = vec![
        ("/_rels/.rels".to_owned(), root_relationships()),
        ("/custom/main.xml".to_owned(), main.to_owned()),
    ];
    if dependencies {
        parts.extend([
            (
                "/custom/_rels/main.xml.rels".to_owned(),
                dependency_relationships(),
            ),
            ("/shared/styles.xml".to_owned(), styles()),
            ("/shared/numbering.xml".to_owned(), numbering()),
        ]);
    }
    parts
}

fn flat_opc(parts: &[(String, String)]) -> String {
    let mut package = format!(r#"<pkg:package xmlns:pkg="{PACKAGE_NAMESPACE}">"#);
    for (name, xml) in parts {
        write!(&mut package, r#"<pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>"#).expect("fixture XML should write");
    }
    package.push_str("</pkg:package>");
    package
}

fn zipped(parts: &[(String, String)]) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, xml) in parts {
        writer
            .start_file(
                name.trim_start_matches('/'),
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )
            .expect("fixture ZIP entry should start");
        writer
            .write_all(xml.as_bytes())
            .expect("fixture ZIP entry should write");
    }
    writer
        .finish()
        .expect("fixture ZIP should finish")
        .into_inner()
}

fn assert_partial(projection: &DocumentProjection) {
    let facts = &projection.structural_facts;
    assert_eq!(
        facts.indentation,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::ParagraphFragment)
    );
    assert_eq!(
        facts.numbering_hierarchy,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::ParagraphFragment)
    );
    assert_eq!(
        facts.bookmarks,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::ParagraphFragment)
    );
    assert_eq!(
        facts.internal_references,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::ParagraphFragment)
    );
    assert_eq!(
        facts.outline_levels,
        StructuralFactSet::Unknown(StructuralFactUnknownReason::ParagraphFragment)
    );
    assert_eq!(projection.paragraphs.len(), 1);
    assert!(
        projection
            .paragraphs
            .iter()
            .all(|paragraph| paragraph.structure.is_none())
    );
}

#[test]
fn text_and_formatting_match_the_same_paragraph_in_a_complete_package() {
    let paragraph = r#"<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>Inherited 😀 </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>direct</w:t></w:r></w:p>"#;
    let fragment = flat_opc(&parts(&document(paragraph), true));
    let full = document(&format!(
        "<w:p><w:r><w:t>Before</w:t></w:r></w:p>{paragraph}<w:p/>"
    ));
    let expected = project_docx(
        &zipped(&parts(&full, true)),
        DocxLimits::default(),
        allocate,
    )
    .expect("complete package should project");
    let actual = project_paragraph_fragment(fragment.as_bytes(), DocxLimits::default(), allocate)
        .expect("fragment should project");
    let full_paragraph = expected
        .paragraphs
        .get(1)
        .expect("complete package target paragraph");
    let fragment_paragraph = actual
        .paragraphs
        .first()
        .expect("fragment target paragraph");
    assert_eq!(fragment_paragraph.text, full_paragraph.text);
    assert_eq!(fragment_paragraph.formatting, full_paragraph.formatting);
    assert_eq!(fragment_paragraph.style_id, full_paragraph.style_id);
    assert!(!fragment_paragraph.formatting.is_empty());
    assert_eq!(
        actual.formatting_status,
        FormattingProjectionStatus::Complete
    );
    assert_eq!(actual.formatting_status, expected.formatting_status);
    assert_partial(&actual);
}

#[test]
fn missing_styles_leave_formatting_incomplete_but_preserve_direct_facts() {
    let main = document(
        r#"<w:p><w:r><w:rPr><w:b/><w:highlight w:val="yellow"/></w:rPr><w:t>Direct</w:t></w:r></w:p>"#,
    );
    let entries = parts(&main, false);
    let expected = project_docx(&zipped(&entries), DocxLimits::default(), allocate)
        .expect("package without styles should project");
    let actual = project_paragraph_fragment(
        flat_opc(&entries).as_bytes(),
        DocxLimits::default(),
        allocate,
    )
    .expect("fragment without styles should project");
    assert_eq!(
        actual.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::StylesPartUnavailable)
    );
    assert_eq!(
        actual
            .paragraphs
            .first()
            .expect("fragment paragraph")
            .formatting,
        expected
            .paragraphs
            .first()
            .expect("complete paragraph")
            .formatting
    );
    assert!(
        !actual
            .paragraphs
            .first()
            .expect("direct paragraph")
            .formatting
            .is_empty()
    );
    assert_partial(&actual);
}

#[test]
fn inherited_package_namespace_aliases_preserve_projection() {
    let source = flat_opc(&parts(
        &document("<w:p><w:r><w:t>Aliases</w:t></w:r></w:p>"),
        false,
    ));
    let aliased = source
        .replace("pkg:", "bundle:")
        .replace("xmlns:pkg=", "xmlns:bundle=")
        .replace("w:", "text:")
        .replace(&format!("xmlns:w=\"{WORD_NAMESPACE}\""), "")
        .replace(
            &format!("xmlns:bundle=\"{PACKAGE_NAMESPACE}\""),
            &format!("xmlns:bundle=\"{PACKAGE_NAMESPACE}\" xmlns:text=\"{WORD_NAMESPACE}\""),
        );
    assert_eq!(
        project_paragraph_fragment(aliased.as_bytes(), DocxLimits::default(), allocate),
        project_paragraph_fragment(source.as_bytes(), DocxLimits::default(), allocate)
    );
}

#[test]
fn all_structural_families_are_unknown_for_empty_and_populated_evidence() {
    for paragraph in [
        "<w:p/>",
        r#"<w:p><w:pPr><w:ind w:left="720"/><w:outlineLvl w:val="1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr><w:bookmarkStart w:id="1" w:name="mark"/><w:hyperlink w:anchor="mark"><w:r><w:t>Target</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="1"/></w:p>"#,
    ] {
        if paragraph != "<w:p/>" {
            let parent = r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Parent</w:t></w:r></w:p>"#;
            let complete = document(&format!("{parent}{paragraph}"));
            let full = project_docx(
                &zipped(&parts(&complete, true)),
                DocxLimits::default(),
                allocate,
            )
            .expect("populated complete structural evidence");
            assert!(
                matches!(&full.structural_facts.indentation, StructuralFactSet::Known(items) if !items.is_empty())
            );
            assert!(
                matches!(&full.structural_facts.numbering_hierarchy, StructuralFactSet::Known(items) if !items.is_empty())
            );
            assert!(
                matches!(&full.structural_facts.bookmarks, StructuralFactSet::Known(items) if !items.is_empty())
            );
            assert!(
                matches!(&full.structural_facts.internal_references, StructuralFactSet::Known(items) if !items.is_empty())
            );
            assert!(
                matches!(&full.structural_facts.outline_levels, StructuralFactSet::Known(items) if !items.is_empty())
            );
        }
        let fragment = flat_opc(&parts(&document(paragraph), true));
        let actual =
            project_paragraph_fragment(fragment.as_bytes(), DocxLimits::default(), allocate)
                .expect("partial structural evidence should project");
        assert_partial(&actual);
    }
}

#[test]
fn table_coordinates_from_a_partial_package_are_not_published() {
    let main =
        document("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
    let entries = parts(&main, true);
    let full = project_docx(&zipped(&entries), DocxLimits::default(), allocate)
        .expect("complete table should project");
    assert!(
        full.paragraphs
            .first()
            .expect("complete cell paragraph")
            .structure
            .is_some()
    );
    let fragment = project_paragraph_fragment(
        flat_opc(&entries).as_bytes(),
        DocxLimits::default(),
        allocate,
    )
    .expect("partial table should project");
    assert_partial(&fragment);
}

#[test]
fn only_a_flat_package_with_exactly_one_paragraph_is_accepted() {
    assert_eq!(
        project_paragraph_fragment(
            document("<w:p/>").as_bytes(),
            DocxLimits::default(),
            allocate
        ),
        Err(ProjectionError::InvalidFlatOpcPackage)
    );
    for (body, error) in [
        ("", ProjectionError::InvalidDocumentXml),
        ("<w:p/><w:p/>", ProjectionError::TooManyParagraphs),
    ] {
        let fragment = flat_opc(&parts(&document(body), false));
        let mut calls = 0_usize;
        let result = project_paragraph_fragment(
            fragment.as_bytes(),
            DocxLimits {
                maximum_paragraphs: 100,
                ..DocxLimits::default()
            },
            |facts| {
                calls = calls.checked_add(1).expect("bounded allocator calls");
                allocate(facts)
            },
        );
        assert_eq!(result, Err(error));
        assert_eq!(calls, 0);
    }
}

#[test]
fn shared_input_dependency_and_paragraph_bounds_are_enforced() {
    let fragment = flat_opc(&parts(&document("<w:p/>"), true));
    for (limits, error) in [
        (
            DocxLimits {
                maximum_archive_bytes: 1,
                ..DocxLimits::default()
            },
            ProjectionError::ArchiveTooLarge,
        ),
        (
            DocxLimits {
                maximum_entries: 1,
                ..DocxLimits::default()
            },
            ProjectionError::TooManyArchiveEntries,
        ),
        (
            DocxLimits {
                maximum_document_xml_bytes: 1,
                ..DocxLimits::default()
            },
            ProjectionError::DocumentXmlTooLarge,
        ),
        (
            DocxLimits {
                maximum_styles_xml_bytes: 1,
                ..DocxLimits::default()
            },
            ProjectionError::StylesXmlTooLarge,
        ),
        (
            DocxLimits {
                maximum_numbering_xml_bytes: 1,
                ..DocxLimits::default()
            },
            ProjectionError::NumberingXmlTooLarge,
        ),
        (
            DocxLimits {
                maximum_numbering_items: 0,
                ..DocxLimits::default()
            },
            ProjectionError::TooManyNumberingItems,
        ),
        (
            DocxLimits {
                maximum_paragraphs: 0,
                ..DocxLimits::default()
            },
            ProjectionError::TooManyParagraphs,
        ),
    ] {
        assert_eq!(
            project_paragraph_fragment(fragment.as_bytes(), limits, allocate),
            Err(error)
        );
    }
    let limited_styles = project_paragraph_fragment(
        fragment.as_bytes(),
        DocxLimits {
            maximum_styles: 0,
            ..DocxLimits::default()
        },
        allocate,
    )
    .expect("unsupported styles preserve direct projection");
    assert_eq!(
        limited_styles.formatting_status,
        FormattingProjectionStatus::Incomplete(FormattingUnknownReason::UnsupportedStyles)
    );
    assert_partial(&limited_styles);
}

#[test]
fn structural_bounds_and_allocator_failures_cannot_be_hidden_by_unknown_output() {
    let main = document(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="first"/><w:bookmarkStart w:id="2" w:name="second"/><w:r><w:t>Content</w:t></w:r><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="1"/></w:p>"#,
    );
    let fragment = flat_opc(&parts(&main, true));
    assert_eq!(
        project_paragraph_fragment(
            fragment.as_bytes(),
            DocxLimits {
                maximum_structural_facts: 1,
                ..DocxLimits::default()
            },
            allocate
        ),
        Err(ProjectionError::TooManyStructuralFacts)
    );
    assert_eq!(
        project_paragraph_fragment(fragment.as_bytes(), DocxLimits::default(), |_| Err(
            ProjectionError::InvalidInternalParagraphId
        )),
        Err(ProjectionError::InvalidInternalParagraphId)
    );
    let mut observed = Vec::new();
    let actual = project_paragraph_fragment(fragment.as_bytes(), DocxLimits::default(), |facts| {
        observed.push((facts.ordinal, facts.text.to_owned()));
        InternalParagraphId::new("allocated-fragment")
    })
    .expect("valid allocator should run once");
    assert_eq!(observed, [(0, "Content".to_owned())]);
    assert_eq!(
        actual
            .paragraphs
            .first()
            .expect("allocated paragraph")
            .id
            .as_str(),
        "allocated-fragment"
    );
    assert_partial(&actual);
}
