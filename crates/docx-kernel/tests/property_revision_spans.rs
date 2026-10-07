#![allow(clippy::expect_used, clippy::panic, clippy::unwrap_used)]

use std::collections::BTreeSet;
use std::io::{Cursor, Write};

use proptest::prelude::{prop_assert_eq, proptest};
use stella_docx_kernel::{
    DocxLimits, InternalParagraphId, ProjectionOptions, ReviewDetail, ReviewFactLimits,
    ReviewFactSet, ReviewPoint, ReviewSpan, RevisionContent, RevisionFactKind, RevisionPayload, RevisionView,
    project_docx_with_review_facts,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

proptest! {
    #[test]
    fn property_changes_cover_their_owner(
        prefix in "[a-z]{0,8}",
        run_text in proptest::sample::select(vec!["", "ab", "é", "😀", "a😀é"]),
        suffix in "[a-z]{0,8}",
        deleted in proptest::bool::ANY,
    ) {
        for view in [RevisionView::Current, RevisionView::Original] {
            let wrapper = if deleted { "del" } else { "ins" };
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:rPr><w:rPrChange w:id="mark" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:pPrChange w:id="paragraph" w:author="A"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>{prefix}</w:t></w:r><w:{wrapper} w:id="text" w:author="A"><w:r><w:rPr><w:rPrChange w:id="run" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>{run_text}</w:t></w:r></w:{wrapper}><w:r><w:t>{suffix}</w:t></w:r></w:p></w:body></w:document>"#
            );
            let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
            archive.start_file("word/document.xml", SimpleFileOptions::default()).unwrap();
            archive.write_all(xml.as_bytes()).unwrap();
            let package = archive.finish().unwrap().into_inner();
            let projection = project_docx_with_review_facts(
                &package,
                DocxLimits::default(),
                ReviewFactLimits::default(),
                ProjectionOptions { revision_view: view, ..ProjectionOptions::default() },
                |facts| InternalParagraphId::new(format!("paragraph-{}", facts.ordinal)),
            ).unwrap();
            let ReviewFactSet::Known(revisions) = projection.review_facts.revisions else {
                panic!("bounded revisions must stay known");
            };
            let property_ids = revisions.iter()
                .filter(|revision| matches!(revision.kind,
                    RevisionFactKind::RunPropertiesChange | RevisionFactKind::ParagraphPropertiesChange))
                .map(|revision| revision.revision_id.as_deref())
                .collect::<Vec<_>>();
            prop_assert_eq!(property_ids.len(), 3, "every property owner has one revision");
            prop_assert_eq!(property_ids.into_iter().collect::<BTreeSet<_>>(),
                BTreeSet::from([Some("paragraph"), Some("mark"), Some("run")]));
            let hidden = matches!((deleted, view), (true, RevisionView::Current) | (false, RevisionView::Original));
            let visible_run = if hidden { "" } else { run_text };
            let paragraph = format!("{prefix}{visible_run}{suffix}");
            let point = |text: &str| ReviewPoint {
                paragraph_ordinal: 0,
                utf8: u32::try_from(text.len()).unwrap(),
                utf16: u32::try_from(text.encode_utf16().count()).unwrap(),
            };
            let end = point(&paragraph);
            for revision in revisions {
                let span = match revision.revision_id.as_deref() {
                    Some("paragraph") => ReviewSpan { start: point(""), end },
                    Some("mark") => ReviewSpan { start: end, end },
                    Some("run") => ReviewSpan { start: point(&prefix), end: point(&format!("{prefix}{visible_run}")) },
                    _ => continue,
                };
                prop_assert_eq!(revision.content, ReviewDetail::Known(RevisionContent {
                    span,
                    payload: RevisionPayload::FormattingOnly,
                }));
            }
        }
    }
}
