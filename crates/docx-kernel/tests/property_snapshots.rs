#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::{prop_assert_eq, proptest};
use stella_docx_kernel::{
    DocumentPackageProjection, DocxLimits, InternalParagraphId, ProjectionOptions,
    ReviewFactLimits, RevisionView, project_docx_with_review_facts,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const NAMESPACES: [&str; 2] = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
];

fn project(
    body: &str,
    styles: &str,
    namespace: &str,
    view: RevisionView,
) -> DocumentPackageProjection {
    let document =
        format!(r#"<a:document xmlns:a="{namespace}"><a:body>{body}</a:body></a:document>"#);
    let style_part = format!(r#"<a:styles xmlns:a="{namespace}">{styles}</a:styles>"#);
    let numbering = format!(
        r#"<a:numbering xmlns:a="{namespace}"><a:abstractNum a:abstractNumId="1"><a:lvl a:ilvl="0"><a:start a:val="1"/><a:numFmt a:val="decimal"/><a:lvlText a:val="%1."/></a:lvl><a:lvl a:ilvl="1"><a:start a:val="1"/><a:numFmt a:val="decimal"/><a:lvlText a:val="%1.%2."/></a:lvl></a:abstractNum><a:num a:numId="1"><a:abstractNumId a:val="1"/></a:num><a:num a:numId="2"><a:abstractNumId a:val="1"/></a:num></a:numbering>"#
    );
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    for (path, xml) in [
        ("word/document.xml", document.as_str()),
        ("word/styles.xml", style_part.as_str()),
        ("word/numbering.xml", numbering.as_str()),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#,
        ),
    ] {
        archive
            .start_file(path, SimpleFileOptions::default())
            .unwrap();
        archive.write_all(xml.as_bytes()).unwrap();
    }
    project_docx_with_review_facts(
        &archive.finish().unwrap().into_inner(),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions {
            revision_view: view,
            ..ProjectionOptions::default()
        },
        |facts| InternalParagraphId::new(format!("paragraph-{}", facts.ordinal)),
    )
    .unwrap()
}

fn run_properties(flags: u8, style: bool) -> String {
    let mut properties = String::new();
    if style {
        properties.push_str(r#"<a:rStyle a:val="C"/>"#);
    }
    for (bit, name) in [(0, "b"), (1, "bCs"), (2, "cs"), (3, "rtl")] {
        let enabled = flags & (1 << bit) != 0;
        if enabled || flags & 128 != 0 {
            write!(properties, r#"<a:{name} a:val="{}"/>"#, u8::from(enabled)).unwrap();
        }
    }
    if flags & 16 != 0 {
        properties.push_str(r#"<a:highlight a:val="yellow"/>"#);
    }
    if flags & 32 != 0 {
        properties.push_str(r#"<a:vertAlign a:val="superscript"/>"#);
    }
    if flags & 64 != 0 {
        properties.push_str("<a:vanish/>");
    }
    properties
}

proptest! {
    #[test]
    fn selected_snapshot_matches_its_direct_property_document(
        before in proptest::num::u8::ANY,
        after in proptest::num::u8::ANY,
        before_style in proptest::bool::ANY,
        after_style in proptest::bool::ANY,
        indent in -2000i32..2000,
        outline in 0u8..10,
        prior_numbering_id in 0u8..3,
        prior_numbering_level in 0u8..2,
        alignment in proptest::sample::select(vec!["left", "right", "center", "both"]),
        cycle in proptest::bool::ANY,
        text in proptest::sample::select(vec!["", "abc", "é😀 العربية"]),
    ) {
        let prior_run = run_properties(before, before_style);
        let current_run = run_properties(after, after_style);
        let prior_paragraph = format!(r#"<a:pStyle a:val="P"/><a:ind a:left="{indent}"/><a:outlineLvl a:val="{outline}"/><a:jc a:val="{alignment}"/><a:numPr><a:numId a:val="{prior_numbering_id}"/><a:ilvl a:val="{prior_numbering_level}"/></a:numPr>"#);
        let current_paragraph = r#"<a:ind a:left="4000"/><a:outlineLvl a:val="8"/><a:jc a:val="center"/><a:numPr><a:numId a:val="1"/><a:ilvl a:val="1"/></a:numPr>"#;
        let styles = format!(r#"<a:style a:type="paragraph" a:styleId="P"><a:basedOn a:val="Q"/><a:rPr><a:b/></a:rPr></a:style><a:style a:type="paragraph" a:styleId="Q">{}<a:pPr><a:ind a:left="120"/></a:pPr></a:style><a:style a:type="character" a:styleId="C"><a:rPr><a:b/><a:highlight a:val="yellow"/></a:rPr></a:style>"#, if cycle { r#"<a:basedOn a:val="P"/>"# } else { "" });
        let child = format!(r#"<a:p><a:pPr><a:numPr><a:numId a:val="{prior_numbering_id}"/><a:ilvl a:val="1"/></a:numPr></a:pPr><a:r><a:t>child</a:t></a:r></a:p>"#);
        let changed = format!(r#"<a:p><a:pPr>{current_paragraph}<a:pPrChange a:id="1" a:author="A"><a:pPr>{prior_paragraph}</a:pPr></a:pPrChange></a:pPr><a:r><a:rPr>{current_run}<a:rPrChange a:id="2" a:author="A"><a:rPr>{prior_run}</a:rPr></a:rPrChange></a:rPr><a:t xml:space="preserve">{text}</a:t></a:r></a:p>{child}"#);
        let prior = format!(r#"<a:p><a:pPr>{prior_paragraph}</a:pPr><a:r><a:rPr>{prior_run}</a:rPr><a:t xml:space="preserve">{text}</a:t></a:r></a:p>{child}"#);
        let current = format!(r#"<a:p><a:pPr>{current_paragraph}</a:pPr><a:r><a:rPr>{current_run}</a:rPr><a:t xml:space="preserve">{text}</a:t></a:r></a:p>{child}"#);
        for namespace in NAMESPACES {
            let restored = project(&changed, &styles, namespace, RevisionView::Original);
            let expected = project(&prior, &styles, namespace, RevisionView::Original);
            if !cycle {
                let stella_docx_kernel::StructuralFactSet::Known(hierarchy) = &restored.document.structural_facts.numbering_hierarchy else {
                    panic!("valid numbering catalogs and noncyclic styles must stay known");
                };
                prop_assert_eq!(hierarchy.len(), if prior_numbering_id != 0 && prior_numbering_level == 0 { 2 } else { 0 });
            }
            prop_assert_eq!(restored.document, expected.document);
            let kept = project(&changed, &styles, namespace, RevisionView::Current);
            let expected_current = project(&current, &styles, namespace, RevisionView::Current);
            prop_assert_eq!(kept.document, expected_current.document);
        }
    }
}

#[test]
fn empty_snapshots_clear_direct_properties_and_keep_table_context() {
    let styles = r#"<a:style a:type="table" a:styleId="T"><a:rPr><a:b/></a:rPr><a:pPr><a:ind a:left="120"/></a:pPr></a:style>"#;
    let changed = r#"<a:tbl><a:tblPr><a:tblStyle a:val="T"/></a:tblPr><a:tr><a:tc><a:p><a:pPr><a:ind a:left="999"/><a:pPrChange a:id="1"><a:pPr/></a:pPrChange></a:pPr><a:r><a:rPr><a:highlight a:val="yellow"/><a:vanish/><a:rPrChange a:id="2"><a:rPr/></a:rPrChange></a:rPr><a:t>x</a:t></a:r></a:p></a:tc></a:tr></a:tbl>"#;
    let expected = r#"<a:tbl><a:tblPr><a:tblStyle a:val="T"/></a:tblPr><a:tr><a:tc><a:p><a:r><a:t>x</a:t></a:r></a:p></a:tc></a:tr></a:tbl>"#;
    for namespace in NAMESPACES {
        assert_eq!(
            project(changed, styles, namespace, RevisionView::Original).document,
            project(expected, styles, namespace, RevisionView::Original).document
        );
    }
}

#[test]
fn table_property_snapshot_restores_selected_style() {
    let styles = r#"<a:style a:type="table" a:styleId="T"><a:rPr><a:b/></a:rPr></a:style><a:style a:type="table" a:styleId="U"><a:rPr><a:highlight a:val="yellow"/></a:rPr></a:style>"#;
    for prior_style in ["", r#"<a:tblStyle a:val="T"/>"#] {
        let changed = format!(
            r#"<a:tbl><a:tblPr><a:tblStyle a:val="U"/><a:tblPrChange a:id="1"><a:tblPr>{prior_style}</a:tblPr></a:tblPrChange></a:tblPr><a:tr><a:tc><a:p><a:r><a:t>x</a:t></a:r></a:p></a:tc></a:tr></a:tbl>"#
        );
        let expected = format!(
            "<a:tbl><a:tblPr>{prior_style}</a:tblPr><a:tr><a:tc><a:p><a:r><a:t>x</a:t></a:r></a:p></a:tc></a:tr></a:tbl>"
        );
        for namespace in NAMESPACES {
            assert_eq!(
                project(&changed, styles, namespace, RevisionView::Original).document,
                project(&expected, styles, namespace, RevisionView::Original).document
            );
        }
    }
}

#[test]
fn historical_metadata_and_ignored_subtrees_do_not_change_facts() {
    let changed = r#"<a:p><a:bookmarkStart a:id="0" a:name="b"/><a:r><a:rPr><a:b/><a:rPrChange a:id="1"><a:rPr><a:unknown><a:vanish/><a:rPrChange a:id="nested"><a:rPr><a:highlight a:val="yellow"/></a:rPr></a:rPrChange><a:commentRangeStart a:id="0"/><a:bookmarkStart a:id="9" a:name="hidden"/></a:unknown></a:rPr></a:rPrChange></a:rPr><a:t>x</a:t></a:r><a:bookmarkEnd a:id="0"/></a:p>"#;
    let expected = r#"<a:p><a:bookmarkStart a:id="0" a:name="b"/><a:r><a:t>x</a:t></a:r><a:bookmarkEnd a:id="0"/></a:p>"#;
    for namespace in NAMESPACES {
        let restored = project(changed, "", namespace, RevisionView::Original);
        assert_eq!(
            restored.document,
            project(expected, "", namespace, RevisionView::Original).document
        );
        let stella_docx_kernel::ReviewFactSet::Known(revisions) = restored.review_facts.revisions
        else {
            panic!("bounded metadata must remain known");
        };
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions.first().unwrap().revision_id.as_deref(), Some("1"));
    }
}

#[test]
fn unavailable_previous_property_roots_remain_incomplete() {
    for snapshot in [
        "",
        "<a:pPr/>",
        r#"<f:rPr xmlns:f="urn:unrelated"/>"#,
        "<a:rPr/><a:rPr/>",
    ] {
        for namespace in NAMESPACES {
            let body = format!(
                r#"<a:p><a:r><a:rPr><a:rPrChange a:id="1">{snapshot}</a:rPrChange></a:rPr><a:t>x</a:t></a:r></a:p>"#
            );
            assert_eq!(
                project(&body, "", namespace, RevisionView::Original)
                    .document
                    .revision_status,
                stella_docx_kernel::RevisionProjectionStatus::Incomplete(vec![
                    stella_docx_kernel::RevisionUnsupportedReason::UnsupportedRevisionMarkup,
                ]),
            );
        }
    }
}

#[test]
fn container_property_snapshots_keep_source_content_and_row_exception_facts() {
    let body = r#"<a:tbl><a:tblPr><a:tblPrChange a:id="table"><a:tblPr/></a:tblPrChange></a:tblPr><a:tblGrid><a:tblGridChange a:id="grid"><a:tblGrid><a:gridCol a:w="120"/></a:tblGrid></a:tblGridChange></a:tblGrid><a:tr><a:tblPrEx><a:tblPrExChange a:id="exception"><a:tblPrEx><a:tblW a:w="120" a:type="dxa"/></a:tblPrEx></a:tblPrExChange></a:tblPrEx><a:trPr><a:trPrChange a:id="row"><a:trPr><a:trHeight a:val="120"/></a:trPr></a:trPrChange></a:trPr><a:tc><a:tcPr><a:tcPrChange a:id="cell"><a:tcPr><a:tcW a:w="120" a:type="dxa"/></a:tcPr></a:tcPrChange></a:tcPr><a:p><a:r><a:t>é😀</a:t></a:r></a:p></a:tc></a:tr></a:tbl><a:sectPr><a:sectPrChange a:id="section"><a:sectPr/></a:sectPrChange></a:sectPr>"#;
    for namespace in NAMESPACES {
        for view in [RevisionView::Current, RevisionView::Original] {
            let projection = project(body, "", namespace, view);
            assert_eq!(
                projection.document.revision_status,
                stella_docx_kernel::RevisionProjectionStatus::Complete
            );
            assert_eq!(projection.document.paragraphs.first().unwrap().text, "é😀");
            let stella_docx_kernel::ReviewFactSet::Known(revisions) =
                projection.review_facts.revisions
            else {
                panic!("bounded review metadata must stay known");
            };
            let exception = revisions
                .iter()
                .find(|revision| revision.revision_id.as_deref() == Some("exception"))
                .unwrap();
            assert_eq!(
                exception.kind,
                stella_docx_kernel::RevisionFactKind::TablePropertiesExceptionChange
            );
            let stella_docx_kernel::ReviewDetail::Known(content) = &exception.content else {
                panic!("row exceptions cover their owning content");
            };
            assert_eq!(content.span.start.paragraph_ordinal, 0);
            assert_eq!(content.span.start.utf8, 0);
            assert_eq!(content.span.end.utf8, 6);
            assert_eq!(content.span.end.utf16, 3);
            assert_eq!(
                content.payload,
                stella_docx_kernel::RevisionPayload::FormattingOnly
            );
        }
    }
}

#[test]
fn paragraph_property_snapshot_preserves_independent_mark_revision() {
    for mark in ["ins", "del"] {
        let changed = format!(
            r#"<a:p><a:pPr><a:ind a:left="120"/><a:rPr><a:{mark} a:id="mark"/></a:rPr><a:pPrChange a:id="properties"><a:pPr/></a:pPrChange></a:pPr><a:r><a:t>a</a:t></a:r></a:p><a:p><a:r><a:t>b</a:t></a:r></a:p>"#
        );
        let expected = format!(
            r#"<a:p><a:pPr><a:rPr><a:{mark} a:id="mark"/></a:rPr></a:pPr><a:r><a:t>a</a:t></a:r></a:p><a:p><a:r><a:t>b</a:t></a:r></a:p>"#
        );
        for namespace in NAMESPACES {
            assert_eq!(
                project(&changed, "", namespace, RevisionView::Original).document,
                project(&expected, "", namespace, RevisionView::Original).document
            );
        }
    }
}

#[test]
fn paragraph_mark_property_snapshot_restores_previous_mark_state() {
    for current in ["", "ins", "del"] {
        for prior in ["", "ins", "del"] {
            let current_marker = if current.is_empty() {
                String::new()
            } else {
                format!(r#"<a:{current} a:id="current"/>"#)
            };
            let prior_marker = if prior.is_empty() {
                String::new()
            } else {
                format!(r#"<a:{prior} a:id="prior"/>"#)
            };
            let changed = format!(
                r#"<a:p><a:pPr><a:rPr>{current_marker}<a:rPrChange a:id="properties"><a:rPr>{prior_marker}</a:rPr></a:rPrChange></a:rPr></a:pPr><a:r><a:t>é</a:t></a:r></a:p><a:p><a:r><a:t>😀</a:t></a:r></a:p>"#
            );
            let expected = format!(
                "<a:p><a:pPr><a:rPr>{prior_marker}</a:rPr></a:pPr><a:r><a:t>é</a:t></a:r></a:p><a:p><a:r><a:t>😀</a:t></a:r></a:p>"
            );
            for namespace in NAMESPACES {
                let restored = project(&changed, "", namespace, RevisionView::Original);
                assert_eq!(
                    restored.document,
                    project(&expected, "", namespace, RevisionView::Original).document
                );
                let stella_docx_kernel::ReviewFactSet::Known(revisions) =
                    restored.review_facts.revisions
                else {
                    panic!("bounded mark metadata must stay known");
                };
                assert!(
                    revisions
                        .iter()
                        .all(|revision| revision.revision_id.as_deref() != Some("prior"))
                );
            }
        }
    }
}

#[test]
fn previous_moved_mark_state_remains_explicitly_incomplete() {
    for prior in ["moveFrom", "moveTo"] {
        let body = format!(
            r#"<a:p><a:pPr><a:rPr><a:rPrChange a:id="properties"><a:rPr><a:{prior} a:id="prior"/></a:rPr></a:rPrChange></a:rPr></a:pPr><a:r><a:t>a</a:t></a:r></a:p><a:p><a:r><a:t>b</a:t></a:r></a:p>"#
        );
        for namespace in NAMESPACES {
            assert_eq!(
                project(&body, "", namespace, RevisionView::Original)
                    .document
                    .revision_status,
                stella_docx_kernel::RevisionProjectionStatus::Incomplete(vec![
                    stella_docx_kernel::RevisionUnsupportedReason::UnsupportedRevisionMarkup
                ])
            );
        }
    }
}

#[test]
fn snapshot_containers_without_their_structural_owner_remain_incomplete() {
    for (properties, change) in [
        ("tblPr", "tblPrChange"),
        ("trPr", "trPrChange"),
        ("tcPr", "tcPrChange"),
        ("tblGrid", "tblGridChange"),
        ("sectPr", "sectPrChange"),
        ("tblPrEx", "tblPrExChange"),
    ] {
        for namespace in NAMESPACES {
            let body = format!(
                r#"<a:p><a:{properties}><a:{change} a:id="1"><a:{properties}/></a:{change}></a:{properties}><a:r><a:t>x</a:t></a:r></a:p>"#
            );
            assert_eq!(
                project(&body, "", namespace, RevisionView::Original)
                    .document
                    .revision_status,
                stella_docx_kernel::RevisionProjectionStatus::Incomplete(vec![
                    stella_docx_kernel::RevisionUnsupportedReason::UnsupportedRevisionMarkup
                ])
            );
        }
    }
}
