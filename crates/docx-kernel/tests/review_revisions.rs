#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used
)]
// Review-revision locations, checked against a structured generator whose
// model computes every expected span and paragraph text independently.

use std::collections::BTreeSet;
use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::{Just, Strategy, any, prop_oneof};
use proptest::{collection, prop_assert_eq, proptest, sample};
use stella_docx_kernel::{
    AttributedRevision, CommentContent, DocumentPackageProjection, DocxLimits, InternalParagraphId,
    ParagraphIdentityFacts, ProjectionError, ProjectionOptions, ReviewDetail, ReviewFactLimits,
    ReviewFactSet, ReviewFactUnknownReason, ReviewPoint, ReviewSpan, RevisionContent,
    RevisionFactKind, RevisionPayload, RevisionProjectionStatus, RevisionUnsupportedReason,
    RevisionView, project_docx_with_review_facts,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const VIEWS: [RevisionView; 2] = [RevisionView::Current, RevisionView::Original];

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("internal-{}", facts.ordinal))
}

fn package(document: &str, comments: Option<&str>) -> Vec<u8> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    archive.start_file("word/document.xml", options).unwrap();
    archive.write_all(document.as_bytes()).unwrap();
    if let Some(comments) = comments {
        archive.start_file("word/comments.xml", options).unwrap();
        archive.write_all(comments.as_bytes()).unwrap();
        archive
            .start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        archive
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>"#,
            )
            .unwrap();
    }
    archive.finish().unwrap().into_inner()
}

fn project(
    document: &str,
    comments: Option<&str>,
    view: RevisionView,
) -> DocumentPackageProjection {
    project_docx_with_review_facts(
        &package(document, comments),
        DocxLimits::default(),
        ReviewFactLimits::default(),
        ProjectionOptions {
            revision_view: view,
            ..ProjectionOptions::default()
        },
        allocate,
    )
    .unwrap()
}

fn known_revisions(projection: &DocumentPackageProjection) -> &[AttributedRevision] {
    let ReviewFactSet::Known(revisions) = &projection.review_facts.revisions else {
        panic!("bounded revision facts should be known");
    };
    revisions
}

const fn point(paragraph_ordinal: usize, utf8: u32, utf16: u32) -> ReviewPoint {
    ReviewPoint {
        paragraph_ordinal,
        utf8,
        utf16,
    }
}

const fn mark_at(at: ReviewPoint) -> ReviewDetail<RevisionContent> {
    ReviewDetail::Known(RevisionContent {
        span: ReviewSpan { start: at, end: at },
        payload: RevisionPayload::ParagraphMark,
    })
}

#[test]
fn paragraph_mark_revisions_are_located_at_the_mark_in_each_view() {
    for (tag, kind, current_texts, original_texts) in [
        (
            "ins",
            RevisionFactKind::Insertion,
            &["A😀", "B"][..],
            &["A😀B"][..],
        ),
        (
            "del",
            RevisionFactKind::Deletion,
            &["A😀B"][..],
            &["A😀", "B"][..],
        ),
        (
            "moveFrom",
            RevisionFactKind::MoveFrom,
            &["A😀B"][..],
            &["A😀", "B"][..],
        ),
        (
            "moveTo",
            RevisionFactKind::MoveTo,
            &["A😀", "B"][..],
            &["A😀B"][..],
        ),
    ] {
        let document = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:rPr><w:{tag} w:id="1" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>A😀</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"#
        );
        for (view, texts) in [
            (RevisionView::Current, current_texts),
            (RevisionView::Original, original_texts),
        ] {
            let projection = project(&document, None, view);
            let projected = projection
                .document
                .paragraphs
                .iter()
                .map(|paragraph| paragraph.text.as_str())
                .collect::<Vec<_>>();
            assert_eq!(projected, texts, "{tag} in {view:?}");
            let revisions = known_revisions(&projection);
            assert_eq!(revisions.len(), 1);
            assert_eq!(revisions[0].kind, kind);
            assert_eq!(
                revisions[0].content,
                mark_at(point(0, 5, 3)),
                "{tag} in {view:?}"
            );
        }
    }
}

#[test]
fn paragraph_merges_relocate_every_review_location() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body>
      <w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>X</w:t></w:r></w:ins></w:p>
      <w:p><w:pPr><w:rPr><w:del w:id="2" w:author="Lin"/></w:rPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>
      <w:p><w:commentRangeStart w:id="9"/><w:ins w:id="3" w:author="Mae"><w:r><w:t>B</w:t></w:r></w:ins><w:commentRangeEnd w:id="9"/></w:p>
    </w:body></w:document>"#
    );
    let comments = format!(
        r#"<w:comments xmlns:w="{W}"><w:comment w:id="9" w:author="Ada"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let projection = project(&document, Some(&comments), RevisionView::Current);
    assert_eq!(projection.document.paragraphs.len(), 2);
    assert_eq!(projection.document.paragraphs[1].text, "AB");
    let contents = known_revisions(&projection)
        .iter()
        .map(|revision| revision.content.clone())
        .collect::<Vec<_>>();
    assert_eq!(
        contents,
        [
            ReviewDetail::Known(RevisionContent {
                span: ReviewSpan {
                    start: point(0, 0, 0),
                    end: point(0, 1, 1),
                },
                payload: RevisionPayload::Text("X".to_owned()),
            }),
            mark_at(point(1, 1, 1)),
            ReviewDetail::Known(RevisionContent {
                span: ReviewSpan {
                    start: point(1, 1, 1),
                    end: point(1, 2, 2),
                },
                payload: RevisionPayload::Text("B".to_owned()),
            }),
        ]
    );
    let ReviewFactSet::Known(projected_comments) = &projection.review_facts.comments else {
        panic!("valid comments should be known");
    };
    assert_eq!(
        projected_comments[0].content,
        ReviewDetail::Known(CommentContent {
            anchor: ReviewSpan {
                start: point(1, 1, 1),
                end: point(1, 2, 2),
            },
            comment_text: "note".to_owned(),
            referenced_text: "B".to_owned(),
        })
    );
}

// ---------------------------------------------------------------------------
// Structured generator and model.

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Tracked {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
}

impl Tracked {
    const fn tag(self) -> &'static str {
        match self {
            Self::Insertion => "ins",
            Self::Deletion => "del",
            Self::MoveFrom => "moveFrom",
            Self::MoveTo => "moveTo",
        }
    }

    const fn kind(self) -> RevisionFactKind {
        match self {
            Self::Insertion => RevisionFactKind::Insertion,
            Self::Deletion => RevisionFactKind::Deletion,
            Self::MoveFrom => RevisionFactKind::MoveFrom,
            Self::MoveTo => RevisionFactKind::MoveTo,
        }
    }

    const fn removes_content(self) -> bool {
        matches!(self, Self::Deletion | Self::MoveFrom)
    }

    const fn hidden_in(self, view: RevisionView) -> bool {
        match view {
            RevisionView::Current => self.removes_content(),
            RevisionView::Original => !self.removes_content(),
        }
    }

    /// A paragraph mark the view hides joins its paragraph to the next one;
    /// Moved marks behave like deleted and inserted ones.
    const fn removes_paragraph_break_in(self, view: RevisionView) -> bool {
        self.hidden_in(view)
    }
}

#[derive(Clone, Debug)]
enum Inline {
    Text(&'static str),
    Revision(Tracked, Vec<Self>),
}

#[derive(Clone, Debug)]
struct Paragraph {
    mark: Option<Tracked>,
    inlines: Vec<Inline>,
    /// Top-level inline range `[start, end)` wrapped by a comment range.
    comment: Option<(usize, usize)>,
    bookmark: bool,
}

#[derive(Clone, Copy, Debug)]
enum TableRevision {
    Row(Tracked),
    CellInsertion,
    CellDeletion,
}

#[derive(Clone, Debug)]
enum Block {
    Paragraph(Paragraph),
    Table {
        revision: Option<TableRevision>,
        cells: Vec<Vec<Paragraph>>,
    },
}

fn tracked() -> impl Strategy<Value = Tracked> {
    sample::select(
        &[
            Tracked::Insertion,
            Tracked::Deletion,
            Tracked::MoveFrom,
            Tracked::MoveTo,
        ][..],
    )
}

fn inline() -> impl Strategy<Value = Inline> {
    let leaf = sample::select(&["a", "é", "😀", " ", "bc"][..]).prop_map(Inline::Text);
    leaf.prop_recursive(4, 24, 3, |inner| {
        (tracked(), collection::vec(inner, 0..3))
            .prop_map(|(kind, children)| Inline::Revision(kind, children))
    })
}

fn paragraph() -> impl Strategy<Value = Paragraph> {
    (
        proptest::option::weighted(0.6, tracked()),
        collection::vec(inline(), 0..4),
        proptest::option::of((any::<u8>(), any::<u8>())),
        any::<bool>(),
    )
        .prop_map(|(mark, inlines, comment, bookmark)| {
            let comment = comment.map(|(start, length)| {
                let start = usize::from(start) % (inlines.len() + 1);
                let end = start + usize::from(length) % (inlines.len() + 1 - start);
                (start, end)
            });
            Paragraph {
                mark,
                inlines,
                comment,
                bookmark,
            }
        })
}

fn block() -> impl Strategy<Value = Block> {
    prop_oneof![
        4 => paragraph().prop_map(Block::Paragraph),
        1 => (
            proptest::option::of(prop_oneof![
                sample::select(&[Tracked::Insertion, Tracked::Deletion][..])
                    .prop_map(TableRevision::Row),
                Just(TableRevision::CellInsertion),
                Just(TableRevision::CellDeletion),
            ]),
            collection::vec(collection::vec(paragraph(), 1..3), 1..3),
        )
            .prop_map(|(revision, cells)| Block::Table { revision, cells }),
    ]
}

struct Markup {
    xml: String,
    next_id: usize,
}

impl Markup {
    fn revision_attributes(&mut self) -> String {
        let id = self.next_id;
        self.next_id += 1;
        format!(r#"w:id="{id}" w:author="Author {}""#, id % 3)
    }

    fn inline(&mut self, inline: &Inline, removed: bool) {
        match inline {
            Inline::Text(text) => {
                let element = if removed { "delText" } else { "t" };
                write!(
                    self.xml,
                    r#"<w:r><w:{element} xml:space="preserve">{text}</w:{element}></w:r>"#
                )
                .unwrap();
            }
            Inline::Revision(kind, children) => {
                let attributes = self.revision_attributes();
                write!(self.xml, "<w:{} {attributes}>", kind.tag()).unwrap();
                for child in children {
                    self.inline(child, removed || kind.removes_content());
                }
                write!(self.xml, "</w:{}>", kind.tag()).unwrap();
            }
        }
    }

    fn paragraph(&mut self, paragraph: &Paragraph, comment_id: usize) {
        self.xml.push_str("<w:p>");
        if let Some(mark) = paragraph.mark {
            let attributes = self.revision_attributes();
            write!(
                self.xml,
                "<w:pPr><w:rPr><w:{} {attributes}/></w:rPr></w:pPr>",
                mark.tag()
            )
            .unwrap();
        }
        for index in 0..=paragraph.inlines.len() {
            if paragraph.comment.is_some_and(|(start, _)| start == index) {
                if paragraph.bookmark {
                    write!(
                        self.xml,
                        r#"<w:bookmarkStart w:id="{comment_id}" w:name="b{comment_id}"/>"#
                    )
                    .unwrap();
                }
                write!(self.xml, r#"<w:commentRangeStart w:id="{comment_id}"/>"#).unwrap();
            }
            if paragraph.comment.is_some_and(|(_, end)| end == index) {
                write!(self.xml, r#"<w:commentRangeEnd w:id="{comment_id}"/>"#).unwrap();
                if paragraph.bookmark {
                    write!(self.xml, r#"<w:bookmarkEnd w:id="{comment_id}"/>"#).unwrap();
                }
            }
            if let Some(inline) = paragraph.inlines.get(index) {
                self.inline(inline, false);
            }
        }
        self.xml.push_str("</w:p>");
    }
}

/// Serializes the model. Comment ids are paragraph source ordinals.
fn write_document(blocks: &[Block]) -> (String, String) {
    let mut markup = Markup {
        xml: format!(r#"<w:document xmlns:w="{W}"><w:body>"#),
        next_id: 1,
    };
    let mut comments = format!(r#"<w:comments xmlns:w="{W}">"#);
    let mut ordinal = 0;
    let mut write_paragraph = |target: &mut Markup, paragraph: &Paragraph| {
        if paragraph.comment.is_some() {
            write!(
                comments,
                r#"<w:comment w:id="{ordinal}" w:author="Ada"><w:p><w:r><w:t>c{ordinal}</w:t></w:r></w:p></w:comment>"#
            )
            .unwrap();
        }
        target.paragraph(paragraph, ordinal);
        ordinal += 1;
    };
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => write_paragraph(&mut markup, paragraph),
            Block::Table { revision, cells } => {
                markup.xml.push_str("<w:tbl><w:tr>");
                if let Some(TableRevision::Row(kind)) = revision {
                    let attributes = markup.revision_attributes();
                    write!(
                        markup.xml,
                        "<w:trPr><w:{} {attributes}/></w:trPr>",
                        kind.tag()
                    )
                    .unwrap();
                }
                for cell in cells {
                    markup.xml.push_str("<w:tc>");
                    let cell_tag = match revision {
                        Some(TableRevision::CellInsertion) => Some("cellIns"),
                        Some(TableRevision::CellDeletion) => Some("cellDel"),
                        _ => None,
                    };
                    if let Some(tag) = cell_tag {
                        let attributes = markup.revision_attributes();
                        write!(markup.xml, "<w:tcPr><w:{tag} {attributes}/></w:tcPr>").unwrap();
                    }
                    for paragraph in cell {
                        write_paragraph(&mut markup, paragraph);
                    }
                    markup.xml.push_str("</w:tc>");
                }
                markup.xml.push_str("</w:tr></w:tbl>");
            }
        }
    }
    markup.xml.push_str("</w:body></w:document>");
    comments.push_str("</w:comments>");
    (markup.xml, comments)
}

/// A point against a source paragraph, before the view joins paragraphs.
#[derive(Clone, Copy, Debug)]
struct SourcePoint {
    paragraph: usize,
    utf8: u32,
    utf16: u32,
}

#[derive(Debug)]
enum ExpectedContent {
    Located {
        start: SourcePoint,
        end: SourcePoint,
        payload: RevisionPayload,
    },
    Unlocated,
}

struct SourceParagraph {
    text: String,
    utf16: u32,
    cell: Option<(usize, usize)>,
    mark: Option<Tracked>,
}

#[derive(Default)]
struct Model {
    view: RevisionView,
    paragraphs: Vec<SourceParagraph>,
    revisions: Vec<(RevisionFactKind, ExpectedContent)>,
    comments: Vec<(usize, SourcePoint, SourcePoint)>,
    structural_table_revision: bool,
}

impl Model {
    fn here(&self) -> SourcePoint {
        let paragraph = self.paragraphs.last().unwrap();
        SourcePoint {
            paragraph: self.paragraphs.len() - 1,
            utf8: u32::try_from(paragraph.text.len()).unwrap(),
            utf16: paragraph.utf16,
        }
    }

    /// Returns the text covered by the inline, visible or not.
    fn inline(&mut self, inline: &Inline, hidden: bool) -> String {
        match inline {
            Inline::Text(text) => {
                if !hidden {
                    let paragraph = self.paragraphs.last_mut().unwrap();
                    paragraph.text.push_str(text);
                    paragraph.utf16 += u32::try_from(text.encode_utf16().count()).unwrap();
                }
                (*text).to_owned()
            }
            Inline::Revision(kind, children) => {
                let index = self.revisions.len();
                self.revisions
                    .push((kind.kind(), ExpectedContent::Unlocated));
                let start = self.here();
                let hidden = hidden || kind.hidden_in(self.view);
                let mut covered = String::new();
                for child in children {
                    covered.push_str(&self.inline(child, hidden));
                }
                let end = self.here();
                self.revisions[index].1 = ExpectedContent::Located {
                    start,
                    end,
                    payload: if covered.is_empty() {
                        RevisionPayload::FormattingOnly
                    } else {
                        RevisionPayload::Text(covered.clone())
                    },
                };
                covered
            }
        }
    }

    fn paragraph(&mut self, paragraph: &Paragraph, cell: Option<(usize, usize)>) {
        let ordinal = self.paragraphs.len();
        self.paragraphs.push(SourceParagraph {
            text: String::new(),
            utf16: 0,
            cell,
            mark: paragraph.mark,
        });
        let mark_index = paragraph.mark.map(|mark| {
            self.revisions
                .push((mark.kind(), ExpectedContent::Unlocated));
            self.revisions.len() - 1
        });
        let mut comment_start = None;
        for index in 0..=paragraph.inlines.len() {
            if let Some((start, end)) = paragraph.comment {
                if start == index {
                    comment_start = Some(self.here());
                }
                if end == index {
                    self.comments
                        .push((ordinal, comment_start.unwrap(), self.here()));
                }
            }
            if let Some(inline) = paragraph.inlines.get(index) {
                self.inline(inline, false);
            }
        }
        if let Some(index) = mark_index {
            let mark = self.here();
            self.revisions[index].1 = ExpectedContent::Located {
                start: mark,
                end: mark,
                payload: RevisionPayload::ParagraphMark,
            };
        }
    }

    fn build(blocks: &[Block], view: RevisionView) -> Self {
        let mut model = Self {
            view,
            ..Self::default()
        };
        for (block_index, block) in blocks.iter().enumerate() {
            match block {
                Block::Paragraph(paragraph) => model.paragraph(paragraph, None),
                Block::Table { revision, cells } => {
                    if let Some(TableRevision::Row(kind)) = revision {
                        model
                            .revisions
                            .push((kind.kind(), ExpectedContent::Unlocated));
                    }
                    model.structural_table_revision |= revision.is_some();
                    for (column, cell) in cells.iter().enumerate() {
                        match revision {
                            Some(TableRevision::CellInsertion) => model.revisions.push((
                                RevisionFactKind::CellInsertion,
                                ExpectedContent::Unlocated,
                            )),
                            Some(TableRevision::CellDeletion) => model
                                .revisions
                                .push((RevisionFactKind::CellDeletion, ExpectedContent::Unlocated)),
                            _ => {}
                        }
                        for paragraph in cell {
                            model.paragraph(paragraph, Some((block_index, column)));
                        }
                    }
                }
            }
        }
        model
    }

    /// Joins paragraphs whose break the view removes, returning the joined
    /// texts, each source paragraph's origin, and whether a join was refused
    /// across a cell boundary.
    fn joined(&self) -> (Vec<String>, Vec<(usize, u32, u32)>, bool) {
        let mut texts: Vec<(String, u32)> = Vec::new();
        let mut origins = Vec::new();
        let mut refused = false;
        for (index, paragraph) in self.paragraphs.iter().enumerate() {
            let joins_previous = index > 0 && {
                let previous = &self.paragraphs[index - 1];
                let removed = previous
                    .mark
                    .is_some_and(|mark| mark.removes_paragraph_break_in(self.view));
                refused |= removed && previous.cell != paragraph.cell;
                removed && previous.cell == paragraph.cell
            };
            if joins_previous {
                let last = texts.len() - 1;
                let (joined, joined_utf16) = &mut texts[last];
                origins.push((last, u32::try_from(joined.len()).unwrap(), *joined_utf16));
                joined.push_str(&paragraph.text);
                *joined_utf16 += paragraph.utf16;
            } else {
                origins.push((texts.len(), 0, 0));
                texts.push((paragraph.text.clone(), paragraph.utf16));
            }
        }
        (
            texts.into_iter().map(|(text, _)| text).collect(),
            origins,
            refused,
        )
    }
}

fn relocate(origins: &[(usize, u32, u32)], at: SourcePoint) -> ReviewPoint {
    let (paragraph_ordinal, utf8, utf16) = origins[at.paragraph];
    point(paragraph_ordinal, utf8 + at.utf8, utf16 + at.utf16)
}

fn slice(texts: &[String], span: ReviewSpan) -> String {
    assert_eq!(
        span.start.paragraph_ordinal, span.end.paragraph_ordinal,
        "generated comment ranges stay inside one paragraph"
    );
    let start = usize::try_from(span.start.utf8).unwrap();
    let end = usize::try_from(span.end.utf8).unwrap();
    texts[span.start.paragraph_ordinal]
        .get(start..end)
        .expect("anchors fall on character boundaries")
        .to_owned()
}

proptest! {
    /// Every inline and paragraph-mark revision is located exactly where the
    /// model puts it in each view, including after paragraph joins; table
    /// row and cell revisions stay explicitly unlocated. Revision text covers
    /// nested revisions; comment anchors survive joins.
    #[test]
    fn generated_tracked_documents_locate_every_revision(
        blocks in collection::vec(block(), 1..8)
    ) {
        let (document, comments) = write_document(&blocks);
        for view in VIEWS {
            let model = Model::build(&blocks, view);
            let (texts, origins, refused) = model.joined();
            let projection = project(&document, Some(&comments), view);
            let projected = projection
                .document
                .paragraphs
                .iter()
                .map(|paragraph| paragraph.text.clone())
                .collect::<Vec<_>>();
            prop_assert_eq!(&projected, &texts);

            let mut reasons = BTreeSet::new();
            if refused {
                reasons.insert(RevisionUnsupportedReason::IncompatibleParagraphMerge);
            }
            if model.structural_table_revision {
                reasons.insert(RevisionUnsupportedReason::StructuralTableRevision);
            }
            let expected_status = if reasons.is_empty() {
                RevisionProjectionStatus::Complete
            } else {
                RevisionProjectionStatus::Incomplete(reasons.into_iter().collect())
            };
            prop_assert_eq!(&projection.document.revision_status, &expected_status);

            let expected = model
                .revisions
                .iter()
                .map(|(kind, content)| {
                    let content = match content {
                        ExpectedContent::Located { start, end, payload } => {
                            ReviewDetail::Known(RevisionContent {
                                span: ReviewSpan {
                                    start: relocate(&origins, *start),
                                    end: relocate(&origins, *end),
                                },
                                payload: payload.clone(),
                            })
                        }
                        ExpectedContent::Unlocated => {
                            ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation)
                        }
                    };
                    (*kind, content)
                })
                .collect::<Vec<_>>();
            let actual = known_revisions(&projection)
                .iter()
                .map(|revision| (revision.kind, revision.content.clone()))
                .collect::<Vec<_>>();
            prop_assert_eq!(&actual, &expected);

            let ReviewFactSet::Known(projected_comments) = &projection.review_facts.comments else {
                panic!("generated comments should be known");
            };
            prop_assert_eq!(projected_comments.len(), model.comments.len());
            for (comment, (id, start, end)) in projected_comments.iter().zip(&model.comments) {
                let anchor = ReviewSpan {
                    start: relocate(&origins, *start),
                    end: relocate(&origins, *end),
                };
                prop_assert_eq!(&comment.comment_id, &id.to_string());
                prop_assert_eq!(
                    &comment.content,
                    &ReviewDetail::Known(CommentContent {
                        anchor,
                        comment_text: format!("c{id}"),
                        referenced_text: slice(&texts, anchor),
                    })
                );
            }
        }
    }
}
