#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]
// Integration fixtures use assertion-style failures; production parsing stays bounded.

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use stella_docx_kernel::{
    DocxLimits, InternalParagraphId, ParagraphIdentityFacts, ProjectionError, StructuralFactSet,
    project_document_xml, project_docx, project_main_document_xml,
};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const FLAT_OPC_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/xmlPackage";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PACKAGE_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("paragraph-{}", facts.ordinal))
}

fn document(body: &str) -> String {
    format!(r#"<w:document xmlns:w="{WORD_NAMESPACE}"><w:body>{body}</w:body></w:document>"#)
}

fn relationships(target: &str) -> String {
    format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="main" Type="{OFFICE_RELATIONSHIPS}/officeDocument" Target="{target}"/></Relationships>"#
    )
}

fn flat_opc(parts: &[(&str, &str)]) -> String {
    let mut package = format!(r#"<pkg:package xmlns:pkg="{FLAT_OPC_NAMESPACE}">"#);
    for (path, xml) in parts {
        write!(&mut package, r#"<pkg:part pkg:name="{path}" pkg:contentType="application/xml"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>"#).expect("test package XML should write");
    }
    package.push_str("</pkg:package>");
    package
}

fn zip_package(parts: &[(&str, &str)]) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (path, xml) in parts {
        writer
            .start_file(
                path.trim_start_matches('/'),
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )
            .expect("test ZIP entry should start");
        writer
            .write_all(xml.as_bytes())
            .expect("test ZIP entry should write");
    }
    writer
        .finish()
        .expect("test ZIP should finish")
        .into_inner()
}

#[test]
fn standalone_document_preserves_the_existing_projection_and_allocator() {
    let xml = document(r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A &amp; 😀</w:t></w:r></w:p><w:p/>"#);
    assert_eq!(
        project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate),
        project_document_xml(xml.as_bytes(), allocate),
    );
    let mut ordinals = Vec::new();
    let projection = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), |facts| {
        ordinals.push(facts.ordinal);
        InternalParagraphId::new(format!("allocated-{}", facts.ordinal))
    })
    .expect("standalone main document should project");
    assert_eq!(ordinals, [0, 1]);
    assert_eq!(projection.paragraphs.len(), ordinals.len());
}

#[test]
fn relationship_selected_main_document_excludes_decoy_bodies_in_every_part_order() {
    let rels = relationships("/custom/main.xml");
    let main = document("<w:p><w:r><w:t>Selected</w:t></w:r></w:p>");
    let decoy = document("<w:p><w:r><w:t>Decoy</w:t></w:r></w:p>");
    let entries = [
        ("/word/document.xml", decoy.as_str()),
        ("/custom/main.xml", main.as_str()),
        ("/_rels/.rels", rels.as_str()),
        ("/word/header1.xml", decoy.as_str()),
    ];
    let expected = project_docx(&zip_package(&entries), DocxLimits::default(), allocate)
        .expect("renamed ZIP main document should project");
    for rotation in 0..entries.len() {
        let mut ordered = entries.to_vec();
        ordered.rotate_left(rotation);
        let projected = project_main_document_xml(
            flat_opc(&ordered).as_bytes(),
            DocxLimits::default(),
            allocate,
        )
        .expect("relationship-selected Flat OPC should project");
        assert_eq!(projected, expected);
        assert_eq!(
            projected
                .paragraphs
                .first()
                .expect("selected paragraph")
                .text,
            "Selected"
        );
    }
}

#[test]
fn strict_and_transitional_packages_preserve_related_styles_and_numbering() {
    for (word_namespace, package_namespace, office_namespace) in [
        (WORD_NAMESPACE, PACKAGE_RELATIONSHIPS, OFFICE_RELATIONSHIPS),
        (
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
            "http://purl.oclc.org/ooxml/package/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ),
    ] {
        let root_rels = format!(
            r#"<Relationships xmlns="{package_namespace}"><Relationship Type="{office_namespace}/officeDocument" Target="/custom/main.xml"/></Relationships>"#
        );
        let doc_rels = format!(
            r#"<Relationships xmlns="{package_namespace}"><Relationship Type="{office_namespace}/styles" Target="../shared/./styles.xml"/><Relationship Type="{office_namespace}/numbering" Target="/lists/numbering.xml"/></Relationships>"#
        );
        let main = format!(
            r#"<w:document xmlns:w="{word_namespace}"><w:body><w:p><w:pPr><w:pStyle w:val="List"/></w:pPr><w:r><w:t>Related</w:t></w:r></w:p></w:body></w:document>"#
        );
        let styles = format!(
            r#"<w:styles xmlns:w="{word_namespace}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>"#
        );
        let numbering = format!(
            r#"<w:numbering xmlns:w="{word_namespace}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num></w:numbering>"#
        );
        let entries = [
            ("/_rels/.rels", root_rels.as_str()),
            ("/custom/main.xml", main.as_str()),
            ("/custom/_rels/main.xml.rels", doc_rels.as_str()),
            ("/shared/styles.xml", styles.as_str()),
            ("/lists/numbering.xml", numbering.as_str()),
        ];
        let expected = project_docx(&zip_package(&entries), DocxLimits::default(), allocate)
            .expect("ZIP with related dependencies should project");
        assert!(
            !expected
                .paragraphs
                .first()
                .expect("related paragraph")
                .formatting
                .is_empty()
        );
        assert!(
            matches!(&expected.structural_facts.indentation, StructuralFactSet::Known(facts) if !facts.is_empty())
        );
        assert_eq!(
            project_main_document_xml(
                flat_opc(&entries).as_bytes(),
                DocxLimits::default(),
                allocate
            ),
            Ok(expected),
        );
    }
}

#[test]
fn namespace_aliases_inherited_from_package_root_survive_part_extraction() {
    for word_namespace in [
        WORD_NAMESPACE,
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let xml = format!(
            r#"<bundle:package xmlns:bundle="{FLAT_OPC_NAMESPACE}" xmlns:text="{word_namespace}" xmlns:rel="{PACKAGE_RELATIONSHIPS}"><bundle:part bundle:name="/_rels/.rels" bundle:contentType="application/xml"><bundle:xmlData><rel:Relationships><rel:Relationship Type="{OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></rel:Relationships></bundle:xmlData></bundle:part><bundle:part bundle:name="/word/document.xml" bundle:contentType="application/xml"><bundle:xmlData><text:document><text:body><text:p><text:r><text:t>Inherited</text:t></text:r></text:p></text:body></text:document></bundle:xmlData></bundle:part></bundle:package>"#
        );
        let projected = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate)
            .expect("inherited aliases should resolve by namespace URI");
        assert_eq!(
            projected
                .paragraphs
                .first()
                .expect("inherited paragraph")
                .text,
            "Inherited"
        );
        let standalone = format!(
            r#"<text:document xmlns:text="{word_namespace}"><text:body><text:p><text:r><text:t>Inherited</text:t></text:r></text:p></text:body></text:document>"#
        );
        assert_eq!(
            projected.paragraphs,
            project_document_xml(standalone.as_bytes(), allocate)
                .expect("standalone alias document")
                .paragraphs
        );
    }
}

#[test]
fn missing_root_relationship_or_selected_part_never_guesses_a_document_path() {
    let main = document("<w:p/>");
    let rels = relationships("missing.xml");
    for entries in [
        vec![("/word/document.xml", main.as_str())],
        vec![
            ("/_rels/.rels", rels.as_str()),
            ("/word/document.xml", main.as_str()),
        ],
    ] {
        assert_eq!(
            project_main_document_xml(
                flat_opc(&entries).as_bytes(),
                DocxLimits::default(),
                allocate
            ),
            Err(ProjectionError::MissingDocumentXml)
        );
    }
}

#[test]
fn duplicate_selected_part_paths_are_rejected() {
    let rels = relationships("word/document.xml");
    let main = document("<w:p/>");
    let xml = flat_opc(&[
        ("/_rels/.rels", &rels),
        ("/word/document.xml", &main),
        ("/word/document.xml", &main),
    ]);
    assert!(matches!(
        project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate),
        Err(ProjectionError::InvalidFlatOpcPackage | ProjectionError::DuplicateDocumentXml)
    ));
}

#[test]
fn main_document_body_is_required_and_unique_for_both_input_forms() {
    let rels = relationships("word/document.xml");
    for (body, expected) in [
        ("", ProjectionError::MissingDocumentBody),
        ("<w:body/><w:body/>", ProjectionError::InvalidDocumentXml),
    ] {
        let main = format!(r#"<w:document xmlns:w="{WORD_NAMESPACE}">{body}</w:document>"#);
        let flat = flat_opc(&[("/_rels/.rels", &rels), ("/word/document.xml", &main)]);
        for input in [main.as_bytes(), flat.as_bytes()] {
            assert_eq!(
                project_main_document_xml(input, DocxLimits::default(), allocate),
                Err(expected.clone())
            );
        }
    }
}

#[test]
fn malformed_package_structure_is_a_typed_failure() {
    let rels = relationships("word/document.xml");
    for xml in [
        format!(
            r#"<pkg:package xmlns:pkg="{FLAT_OPC_NAMESPACE}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml"><pkg:xmlData></pkg:part></pkg:package>"#
        ),
        format!(
            r#"<pkg:package xmlns:pkg="{FLAT_OPC_NAMESPACE}"><pkg:part pkg:contentType="application/xml"><pkg:xmlData>{rels}</pkg:xmlData></pkg:part></pkg:package>"#
        ),
        format!(
            r#"<pkg:package xmlns:pkg="{FLAT_OPC_NAMESPACE}"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/xml"><pkg:xmlData>{rels}</pkg:xmlData><pkg:xmlData>{rels}</pkg:xmlData></pkg:part></pkg:package>"#
        ),
    ] {
        assert_eq!(
            project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate),
            Err(ProjectionError::InvalidFlatOpcPackage)
        );
    }
}

#[test]
fn shared_limits_bound_input_part_count_and_selected_document_bytes() {
    let main = document("<w:p/><w:p/>");
    let rels = relationships("word/document.xml");
    let flat = flat_opc(&[("/_rels/.rels", &rels), ("/word/document.xml", &main)]);
    for input in [main.as_bytes(), flat.as_bytes()] {
        assert_eq!(
            project_main_document_xml(
                input,
                DocxLimits {
                    maximum_archive_bytes: 1,
                    ..DocxLimits::default()
                },
                allocate
            ),
            Err(ProjectionError::ArchiveTooLarge)
        );
        assert_eq!(
            project_main_document_xml(
                input,
                DocxLimits {
                    maximum_document_xml_bytes: 1,
                    ..DocxLimits::default()
                },
                allocate
            ),
            Err(ProjectionError::DocumentXmlTooLarge)
        );
        assert_eq!(
            project_main_document_xml(
                input,
                DocxLimits {
                    maximum_paragraphs: 1,
                    ..DocxLimits::default()
                },
                allocate
            ),
            Err(ProjectionError::TooManyParagraphs)
        );
    }
    assert_eq!(
        project_main_document_xml(
            flat.as_bytes(),
            DocxLimits {
                maximum_entries: 1,
                ..DocxLimits::default()
            },
            allocate
        ),
        Err(ProjectionError::TooManyArchiveEntries)
    );
}

#[test]
fn shared_structural_budget_and_allocator_failures_are_preserved() {
    let main = document(
        r#"<w:p><w:bookmarkStart w:id="1" w:name="first"/><w:bookmarkStart w:id="2" w:name="second"/><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="1"/></w:p>"#,
    );
    let rels = relationships("word/document.xml");
    let flat = flat_opc(&[("/_rels/.rels", &rels), ("/word/document.xml", &main)]);
    for input in [main.as_bytes(), flat.as_bytes()] {
        assert_eq!(
            project_main_document_xml(
                input,
                DocxLimits {
                    maximum_structural_facts: 1,
                    ..DocxLimits::default()
                },
                allocate
            ),
            Err(ProjectionError::TooManyStructuralFacts)
        );
        assert_eq!(
            project_main_document_xml(input, DocxLimits::default(), |_| Err(
                ProjectionError::InvalidInternalParagraphId
            )),
            Err(ProjectionError::InvalidInternalParagraphId)
        );
        assert_eq!(
            project_main_document_xml(input, DocxLimits::default(), |_| InternalParagraphId::new(
                "same"
            )),
            Err(ProjectionError::DuplicateInternalParagraphId)
        );
    }
}

#[test]
fn related_auxiliary_parts_use_the_same_byte_and_item_limits_as_zip() {
    let rels = relationships("word/document.xml");
    let main = document("<w:p/>");
    let doc_rels = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Type="{OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/><Relationship Type="{OFFICE_RELATIONSHIPS}/numbering" Target="numbering.xml"/></Relationships>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{WORD_NAMESPACE}"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/></w:styles>"#
    );
    let numbering = format!(
        r#"<w:numbering xmlns:w="{WORD_NAMESPACE}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"/></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num></w:numbering>"#
    );
    let entries = [
        ("/_rels/.rels", rels.as_str()),
        ("/word/document.xml", main.as_str()),
        ("/word/_rels/document.xml.rels", doc_rels.as_str()),
        ("/word/styles.xml", styles.as_str()),
        ("/word/numbering.xml", numbering.as_str()),
    ];
    let flat = flat_opc(&entries);
    let zipped = zip_package(&entries);
    for (limits, expected) in [
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
    ] {
        assert_eq!(
            project_docx(&zipped, limits, allocate),
            Err(expected.clone())
        );
        assert_eq!(
            project_main_document_xml(flat.as_bytes(), limits, allocate),
            Err(expected)
        );
    }
    let style_limit = DocxLimits {
        maximum_styles: 0,
        ..DocxLimits::default()
    };
    assert_eq!(
        project_main_document_xml(flat.as_bytes(), style_limit, allocate),
        project_docx(&zipped, style_limit, allocate)
    );
}

#[test]
fn raw_input_requires_a_namespace_qualified_document_root() {
    for xml in [
        "<body><p/></body>",
        "<document><body><p/></body></document>",
        "<w:document xmlns:w=\"urn:unrelated\"><w:body/></w:document>",
        "<w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p/></w:body>",
    ] {
        let mut allocator_calls = 0_usize;
        let result = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), |facts| {
            allocator_calls = allocator_calls
                .checked_add(1)
                .expect("bounded allocator call count");
            allocate(facts)
        });
        assert_eq!(result, Err(ProjectionError::InvalidDocumentXml));
        assert_eq!(allocator_calls, 0);
    }
}

#[test]
fn inherited_default_namespace_and_package_paragraph_ids_remain_document_facts() {
    let xml = format!(
        r#"<pkg:package xmlns:pkg="{FLAT_OPC_NAMESPACE}" xmlns="{WORD_NAMESPACE}" xmlns:pid="http://schemas.microsoft.com/office/word/2010/wordml"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/xml"><pkg:xmlData><Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Type="{OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml"><pkg:xmlData><document><body><p pid:paraId="1234ABCD"><r><t>Default namespace</t></r></p></body></document></pkg:xmlData></pkg:part></pkg:package>"#
    );
    let expected = format!(
        r#"<document xmlns="{WORD_NAMESPACE}" xmlns:pid="http://schemas.microsoft.com/office/word/2010/wordml"><body><p pid:paraId="1234ABCD"><r><t>Default namespace</t></r></p></body></document>"#
    );
    let projected = project_main_document_xml(xml.as_bytes(), DocxLimits::default(), allocate)
        .expect("inherited default namespace should project");
    let standalone = project_document_xml(expected.as_bytes(), allocate)
        .expect("standalone default namespace should project");
    assert_eq!(projected.paragraphs, standalone.paragraphs);
    assert!(
        projected
            .paragraphs
            .first()
            .expect("namespace paragraph")
            .package_paragraph_id
            .is_some()
    );
}

#[test]
fn unrelated_zero_byte_binary_parts_preserve_selected_document_projection() {
    let main = document("<w:p><w:r><w:t>Selected</w:t></w:r></w:p>");
    let root_relationships = relationships("word/document.xml");
    let parts = [
        ("/_rels/.rels", root_relationships.as_str()),
        ("/word/document.xml", main.as_str()),
    ];
    let expected =
        project_main_document_xml(flat_opc(&parts).as_bytes(), DocxLimits::default(), allocate)
            .unwrap();
    for payload in ["<pkg:binaryData/>", "<pkg:binaryData></pkg:binaryData>"] {
        let package = flat_opc(&parts).replace("</pkg:package>", &format!("<pkg:part pkg:name=\"/media/empty.bin\" pkg:contentType=\"application/octet-stream\">{payload}</pkg:part></pkg:package>"));
        assert_eq!(
            project_main_document_xml(package.as_bytes(), DocxLimits::default(), allocate).unwrap(),
            expected
        );
    }
}
