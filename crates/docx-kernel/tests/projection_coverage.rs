#![allow(
    clippy::arithmetic_side_effects,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::string_slice,
    clippy::unwrap_used
)]
// Context coverage for the package projection.
//
// A structured generator places every tracked-change and annotation element
// in the supported contexts listed by the suite. Excluded placements are
// named explicitly in `EXCLUDED_ROWS`.
// Four checks keep a placement nobody thought of from passing silently:
//
// - the context table is bound to the parser's own element dispatch, so a new
//   parser branch needs a new row;
// - a coverage receipt proves the generator emits every row;
// - generated facts match an independent semantic model;
// - metamorphic relations compare projections of equivalent documents.
//
// `PROPERTY_TEST_NUM_RUNS_FACTOR` scales every case count (the nightly lane
// sets it); the default keeps PR CI fast.
//
// String slices index the parser source at ASCII delimiters and model text at
// offsets the model computed from whole characters.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::io::{Cursor, Write};

use proptest::prelude::{BoxedStrategy, Just, Strategy, any, prop_oneof};
use proptest::strategy::ValueTree;
use proptest::test_runner::{Config, FileFailurePersistence, TestRunner};
use proptest::{collection, prop_assert, prop_assert_eq, proptest, sample};
use quick_xml::events::{BytesStart, Event};
use stella_docx_kernel::{
    AttributedRevision, CommentContent, DocumentPackageProjection, DocumentStructureFacts,
    DocxLimits, FormattingCompleteness, FormattingFactStatus, InternalParagraphId,
    ParagraphIdentityFacts, ParagraphStructure, ProjectionError, ProjectionOptions, ReviewDetail,
    ReviewFactLimits, ReviewFactSet, ReviewPoint, ReviewSpan, RevisionContent, RevisionFactKind,
    RevisionPayload, RevisionProjectionStatus, RevisionView, StructuralFactSet, TextFormattingSpan,
    TextStyle, project_docx_with_review_facts,
};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const V: &str = "urn:schemas-microsoft-com:vml";
const VIEWS: [RevisionView; 2] = [RevisionView::Current, RevisionView::Original];
const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/></w:styles>"#;
const RELATIONSHIPS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
/// Revision and bookmark id used by metamorphic probes; generated ids stay
/// far below it.
const PROBE_ID: usize = 999_999;
const PROBE_AUTHOR: &str = "Probe";
const RECEIPT_CASES: u32 = 768;

fn case_factor() -> u32 {
    std::env::var("PROPERTY_TEST_NUM_RUNS_FACTOR")
        .ok()
        .and_then(|value| value.parse::<u32>().ok())
        .filter(|factor| *factor > 0)
        .unwrap_or(1)
}

fn config(cases: u32) -> Config {
    Config {
        cases: cases * case_factor(),
        failure_persistence: Some(Box::new(FileFailurePersistence::Direct(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/projection_coverage.proptest-regressions"
        )))),
        ..Config::default()
    }
}

// ---------------------------------------------------------------------------
// Document model.

#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
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

    /// Content the view does not show: removed content in the current view,
    /// added content in the original view.
    const fn hidden_in(self, view: RevisionView) -> bool {
        match view {
            RevisionView::Current => self.removes_content(),
            RevisionView::Original => !self.removes_content(),
        }
    }

    /// ECMA-376 Part 1, 17.13.5: a deleted or moved-away paragraph mark no
    /// longer delimits its paragraph, so its content combines with the next
    /// paragraph; an inserted or moved-here mark did not exist originally.
    const fn removes_paragraph_break_in(self, view: RevisionView) -> bool {
        self.hidden_in(view)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RunContent {
    Text(&'static str),
    Tab,
    Break,
    FootnoteReference,
    EndnoteReference,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Run {
    content: RunContent,
    bold: bool,
    prior_bold: bool,
    property_change: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Author {
    Generated,
    Probe,
}

#[derive(Clone, Debug)]
struct Revision {
    kind: Tracked,
    author: Author,
    children: Vec<Inline>,
}

#[derive(Clone, Debug)]
enum Inline {
    Run(Run),
    Revision(Revision),
    Hyperlink(u8, Vec<Self>),
    Sdt(Vec<Self>),
    SmartTag(Vec<Self>),
    CustomXml(Vec<Self>),
    SimpleField(u8, Vec<Self>),
    ComplexField(u8, Vec<Self>),
    Comment(usize, Vec<Self>),
    Bookmark(usize, Vec<Self>),
    ProofErr,
    Textbox(Vec<Block>),
}

#[derive(Clone, Debug)]
struct Paragraph {
    mark: Option<Tracked>,
    mark_property_change: bool,
    property_change: bool,
    inlines: Vec<Inline>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TableRevision {
    Row(Tracked),
    CellInsertion,
    CellDeletion,
    CellMerge,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TableLevel {
    Table,
    Row,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
struct TablePropertyChanges {
    table: bool,
    row: bool,
    cell: bool,
}

#[derive(Clone, Debug)]
struct Table {
    revision: Option<TableRevision>,
    property_changes: TablePropertyChanges,
    bookmark: Option<(TableLevel, usize)>,
    cells: Vec<Vec<Block>>,
}

#[derive(Clone, Debug)]
enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Sdt(Vec<Self>),
    CustomXml(Vec<Self>),
    Bookmark(usize, Vec<Self>),
    Comment(usize, Vec<Self>),
}

#[derive(Clone, Debug)]
struct Document {
    blocks: Vec<Block>,
    section_change: bool,
}

// ---------------------------------------------------------------------------
// Strategies. Children are generated freely and `sanitize` enforces the
// schema's placement rules, so shrinking keeps documents valid.

const LEAVES: [&str; 6] = ["a", "é", "😀", " ", "bc", "x y"];

fn run() -> impl Strategy<Value = Run> {
    (
        prop_oneof![
            8 => sample::select(&LEAVES[..]).prop_map(RunContent::Text),
            1 => Just(RunContent::Tab),
            1 => Just(RunContent::Break),
            1 => Just(RunContent::FootnoteReference),
            1 => Just(RunContent::EndnoteReference),
        ],
        any::<bool>(),
        any::<bool>(),
        proptest::bool::weighted(0.08),
    )
        .prop_map(|(content, bold, prior_bold, property_change)| Run {
            content,
            bold,
            prior_bold,
            property_change,
        })
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

fn inline(depth: u32) -> BoxedStrategy<Inline> {
    let leaf = prop_oneof![
        10 => run().prop_map(Inline::Run),
        1 => Just(Inline::ProofErr),
    ]
    .boxed();
    if depth == 0 {
        return leaf;
    }
    let children = collection::vec(inline(depth - 1), 0..3);
    let target = 0..4_u8;
    let mut arms = vec![
        (12, leaf),
        (
            6,
            (tracked(), children.clone())
                .prop_map(|(kind, children)| {
                    Inline::Revision(Revision {
                        kind,
                        author: Author::Generated,
                        children,
                    })
                })
                .boxed(),
        ),
        (
            1,
            (target.clone(), children.clone())
                .prop_map(|(target, children)| Inline::Hyperlink(target, children))
                .boxed(),
        ),
        (1, children.clone().prop_map(Inline::Sdt).boxed()),
        (1, children.clone().prop_map(Inline::SmartTag).boxed()),
        (1, children.clone().prop_map(Inline::CustomXml).boxed()),
        (
            1,
            (target.clone(), children.clone())
                .prop_map(|(target, children)| Inline::SimpleField(target, children))
                .boxed(),
        ),
        (
            1,
            (target, children.clone())
                .prop_map(|(target, children)| Inline::ComplexField(target, children))
                .boxed(),
        ),
        (
            2,
            children
                .clone()
                .prop_map(|children| Inline::Comment(0, children))
                .boxed(),
        ),
        (
            2,
            children
                .prop_map(|children| Inline::Bookmark(0, children))
                .boxed(),
        ),
    ];
    if depth >= 2 {
        arms.push((
            1,
            collection::vec(block(1, depth - 2), 1..3)
                .prop_map(Inline::Textbox)
                .boxed(),
        ));
    }
    proptest::strategy::Union::new_weighted(arms).boxed()
}

fn paragraph(inline_depth: u32) -> impl Strategy<Value = Paragraph> {
    (
        proptest::option::weighted(0.3, tracked()),
        proptest::bool::weighted(0.08),
        proptest::bool::weighted(0.08),
        collection::vec(inline(inline_depth), 0..4),
    )
        .prop_map(
            |(mark, mark_property_change, property_change, inlines)| Paragraph {
                mark,
                mark_property_change,
                property_change,
                inlines,
            },
        )
}

fn table_revision() -> impl Strategy<Value = TableRevision> {
    prop_oneof![
        sample::select(&[Tracked::Insertion, Tracked::Deletion][..]).prop_map(TableRevision::Row),
        Just(TableRevision::CellInsertion),
        Just(TableRevision::CellDeletion),
        Just(TableRevision::CellMerge),
    ]
}

fn block(block_depth: u32, inline_depth: u32) -> BoxedStrategy<Block> {
    let paragraph = paragraph(inline_depth).prop_map(Block::Paragraph).boxed();
    if block_depth == 0 {
        return paragraph;
    }
    let children = collection::vec(block(block_depth - 1, inline_depth), 1..3);
    let table = (
        proptest::option::weighted(0.4, table_revision()),
        (
            proptest::bool::weighted(0.1),
            proptest::bool::weighted(0.1),
            proptest::bool::weighted(0.1),
        ),
        proptest::option::weighted(
            0.2,
            sample::select(&[TableLevel::Table, TableLevel::Row][..]),
        ),
        collection::vec(children.clone(), 1..3),
    )
        .prop_map(|(revision, (table, row, cell), bookmark, cells)| {
            Block::Table(Table {
                revision,
                property_changes: TablePropertyChanges { table, row, cell },
                bookmark: bookmark.map(|level| (level, 0)),
                cells,
            })
        });
    prop_oneof![
        8 => paragraph,
        3 => table,
        1 => children.clone().prop_map(Block::Sdt),
        1 => children.clone().prop_map(Block::CustomXml),
        1 => children.clone().prop_map(|children| Block::Bookmark(0, children)),
        1 => children.prop_map(|children| Block::Comment(0, children)),
    ]
    .boxed()
}

fn document() -> impl Strategy<Value = Document> {
    (
        collection::vec(block(2, 3), 1..6),
        proptest::bool::weighted(0.1),
        proptest::option::weighted(0.3, tracked()),
    )
        .prop_map(|(blocks, section_change, terminal_mark)| {
            let mut document = Document {
                blocks,
                section_change,
            };
            sanitize_document(&mut document);
            for index in 0..document.blocks.len() {
                let eligible = index + 1 == document.blocks.len()
                    || matches!(document.blocks.get(index + 1), Some(Block::Paragraph(_)));
                if eligible && let Block::Paragraph(paragraph) = &mut document.blocks[index] {
                    paragraph.mark = terminal_mark;
                }
            }
            document
        })
}

/// Property snapshots are generated for both revision views.
fn property_document() -> impl Strategy<Value = Document> {
    (document(), any::<bool>(), any::<bool>(), any::<bool>()).prop_map(
        |(mut document, paragraph_change, mark_change, run_change)| {
            map_paragraphs(&mut document.blocks, &mut |paragraph| {
                paragraph.property_change = paragraph_change;
                paragraph.mark_property_change = mark_change;
            });
            visit_inline_lists(&mut document.blocks, &mut |inlines| {
                for inline in inlines {
                    if let Inline::Run(run) = inline {
                        run.property_change = run_change;
                    }
                }
            });
            document
        },
    )
}

/// Scopes placements to the suite, then numbers comment
/// and bookmark ids in document order.
fn sanitize_document(document: &mut Document) {
    scope_document(document);
    let mut ids = Ids::default();
    number_blocks(&mut document.blocks, &mut ids);
}

/// Paragraph marks occur on direct body paragraphs followed by a paragraph,
/// or on the final direct body paragraph.
fn scope_document(document: &mut Document) {
    let marks = document
        .blocks
        .iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => paragraph.mark,
            _ => None,
        })
        .collect::<Vec<_>>();
    document.section_change = false;
    sanitize_blocks(&mut document.blocks);
    end_with_paragraph(&mut document.blocks);
    for index in 0..document.blocks.len() {
        let eligible = index + 1 == document.blocks.len()
            || matches!(document.blocks.get(index + 1), Some(Block::Paragraph(_)));
        if eligible && let Block::Paragraph(paragraph) = &mut document.blocks[index] {
            paragraph.mark = marks.get(index).copied().flatten();
        }
    }
}

fn sanitize_blocks(blocks: &mut Vec<Block>) {
    let mut output = Vec::with_capacity(blocks.len());
    for mut block in std::mem::take(blocks) {
        match &mut block {
            Block::Paragraph(paragraph) => {
                paragraph.mark = None;
                paragraph.mark_property_change = false;
                paragraph.property_change = false;
                sanitize_inlines(&mut paragraph.inlines, InlineScope::OutsideRevision);
            }
            Block::Table(table) => {
                for cell in &mut table.cells {
                    sanitize_blocks(cell);
                    output.append(cell);
                }
                continue;
            }
            Block::Sdt(children)
            | Block::CustomXml(children)
            | Block::Bookmark(_, children)
            | Block::Comment(_, children) => sanitize_blocks(children),
        }
        output.push(block);
    }
    *blocks = output;
}

#[derive(Clone, Copy)]
enum InlineScope {
    OutsideRevision,
    Revision,
    RevisionContainer,
}

impl InlineScope {
    const fn inside_revision(self) -> bool {
        !matches!(self, Self::OutsideRevision)
    }
}

fn sanitize_inlines(inlines: &mut Vec<Inline>, scope: InlineScope) {
    let mut output = Vec::with_capacity(inlines.len());
    for mut inline in std::mem::take(inlines) {
        match &mut inline {
            Inline::Run(run) => {
                run.property_change = false;
            }
            Inline::ProofErr => {}
            Inline::Revision(revision) => {
                sanitize_inlines(&mut revision.children, InlineScope::Revision);
            }
            Inline::Hyperlink(_, children) | Inline::SimpleField(_, children) => {
                sanitize_inlines(children, scope);
                if scope.inside_revision() {
                    output.append(children);
                    continue;
                }
            }
            Inline::ComplexField(_, children) | Inline::Bookmark(_, children) => {
                sanitize_inlines(children, scope);
            }
            Inline::Sdt(children) | Inline::SmartTag(children) | Inline::CustomXml(children) => {
                let child_scope = if scope.inside_revision() {
                    InlineScope::RevisionContainer
                } else {
                    InlineScope::OutsideRevision
                };
                sanitize_inlines(children, child_scope);
            }
            Inline::Comment(_, children) => sanitize_inlines(children, scope),
            Inline::Textbox(_) => continue,
        }
        output.push(inline);
    }
    *inlines = output;
}

const fn empty_paragraph() -> Paragraph {
    Paragraph {
        mark: None,
        mark_property_change: false,
        property_change: false,
        inlines: Vec::new(),
    }
}

/// Ensures a nonempty document after placements are scoped.
fn end_with_paragraph(blocks: &mut Vec<Block>) {
    if !matches!(blocks.last(), Some(Block::Paragraph(_))) {
        blocks.push(Block::Paragraph(empty_paragraph()));
    }
}

#[derive(Default)]
struct Ids {
    comment: usize,
    bookmark: usize,
}

impl Ids {
    const fn comment(&mut self) -> usize {
        self.comment += 1;
        self.comment - 1
    }

    const fn bookmark(&mut self) -> usize {
        self.bookmark += 1;
        self.bookmark - 1
    }
}

fn number_blocks(blocks: &mut [Block], ids: &mut Ids) {
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => number_inlines(&mut paragraph.inlines, ids),
            Block::Table(table) => {
                if let Some((_, id)) = &mut table.bookmark {
                    *id = ids.bookmark();
                }
                for cell in &mut table.cells {
                    number_blocks(cell, ids);
                }
            }
            Block::Sdt(children) | Block::CustomXml(children) => number_blocks(children, ids),
            Block::Bookmark(id, children) => {
                *id = ids.bookmark();
                number_blocks(children, ids);
            }
            Block::Comment(id, children) => {
                *id = ids.comment();
                number_blocks(children, ids);
            }
        }
    }
}

fn number_inlines(inlines: &mut [Inline], ids: &mut Ids) {
    for inline in inlines {
        match inline {
            Inline::Run(_) | Inline::ProofErr => {}
            Inline::Revision(revision) => number_inlines(&mut revision.children, ids),
            Inline::Hyperlink(_, children)
            | Inline::Sdt(children)
            | Inline::SmartTag(children)
            | Inline::CustomXml(children)
            | Inline::SimpleField(_, children)
            | Inline::ComplexField(_, children) => number_inlines(children, ids),
            Inline::Comment(id, children) => {
                *id = ids.comment();
                number_inlines(children, ids);
            }
            Inline::Bookmark(id, children) => {
                *id = ids.bookmark();
                number_inlines(children, ids);
            }
            Inline::Textbox(blocks) => number_blocks(blocks, ids),
        }
    }
}

// ---------------------------------------------------------------------------
// Serialization.

struct Markup {
    xml: String,
    comment_ids: BTreeSet<usize>,
    next_revision: usize,
}

impl Markup {
    fn revision_attributes(&mut self, author: Author) -> String {
        match author {
            Author::Probe => format!(r#"w:id="{PROBE_ID}" w:author="{PROBE_AUTHOR}""#),
            Author::Generated => {
                let id = self.next_revision;
                self.next_revision += 1;
                format!(
                    r#"w:id="{id}" w:author="Author {}" w:date="2024-01-0{}T00:00:00Z""#,
                    id % 3,
                    id % 9 + 1
                )
            }
        }
    }

    fn change(&mut self, element: &str, snapshot: &str) {
        let attributes = self.revision_attributes(Author::Generated);
        write!(
            self.xml,
            "<w:{element} {attributes}>{snapshot}</w:{element}>"
        )
        .unwrap();
    }

    fn bookmark_start(&mut self, id: usize) {
        write!(self.xml, r#"<w:bookmarkStart w:id="{id}" w:name="b{id}"/>"#).unwrap();
    }

    fn bookmark_end(&mut self, id: usize) {
        write!(self.xml, r#"<w:bookmarkEnd w:id="{id}"/>"#).unwrap();
    }

    fn run(&mut self, run: Run, removed: bool) {
        self.xml.push_str("<w:r>");
        if run.bold || run.property_change {
            self.xml.push_str("<w:rPr>");
            if run.bold {
                self.xml.push_str("<w:b/>");
            }
            if run.property_change {
                let prior = if run.prior_bold {
                    "<w:rPr><w:b/></w:rPr>"
                } else {
                    "<w:rPr/>"
                };
                self.change("rPrChange", prior);
            }
            self.xml.push_str("</w:rPr>");
        }
        match run.content {
            RunContent::Text(text) => {
                let element = if removed { "delText" } else { "t" };
                write!(
                    self.xml,
                    r#"<w:{element} xml:space="preserve">{text}</w:{element}>"#
                )
                .unwrap();
            }
            RunContent::Tab => self.xml.push_str("<w:tab/>"),
            RunContent::Break => self.xml.push_str("<w:br/>"),
            RunContent::FootnoteReference => {
                self.xml.push_str(r#"<w:footnoteReference w:id="1"/>"#);
            }
            RunContent::EndnoteReference => {
                self.xml.push_str(r#"<w:endnoteReference w:id="1"/>"#);
            }
        }
        self.xml.push_str("</w:r>");
    }

    fn inlines(&mut self, inlines: &[Inline], removed: bool) {
        for inline in inlines {
            self.inline(inline, removed);
        }
    }

    fn inline(&mut self, inline: &Inline, removed: bool) {
        match inline {
            Inline::Run(run) => self.run(*run, removed),
            Inline::Revision(revision) => {
                let attributes = self.revision_attributes(revision.author);
                let tag = revision.kind.tag();
                write!(self.xml, "<w:{tag} {attributes}>").unwrap();
                self.inlines(
                    &revision.children,
                    removed || revision.kind.removes_content(),
                );
                write!(self.xml, "</w:{tag}>").unwrap();
            }
            Inline::Hyperlink(target, children) => {
                write!(self.xml, r#"<w:hyperlink w:anchor="b{target}">"#).unwrap();
                self.inlines(children, removed);
                self.xml.push_str("</w:hyperlink>");
            }
            Inline::Sdt(children) => {
                self.xml.push_str("<w:sdt><w:sdtPr/><w:sdtContent>");
                self.inlines(children, removed);
                self.xml.push_str("</w:sdtContent></w:sdt>");
            }
            Inline::SmartTag(children) => {
                self.xml
                    .push_str(r#"<w:smartTag w:uri="urn:example" w:element="tag">"#);
                self.inlines(children, removed);
                self.xml.push_str("</w:smartTag>");
            }
            Inline::CustomXml(children) => {
                self.xml.push_str(r#"<w:customXml w:element="item">"#);
                self.inlines(children, removed);
                self.xml.push_str("</w:customXml>");
            }
            Inline::SimpleField(target, children) => {
                write!(self.xml, r#"<w:fldSimple w:instr=" REF b{target} \h ">"#).unwrap();
                self.inlines(children, removed);
                self.xml.push_str("</w:fldSimple>");
            }
            Inline::ComplexField(target, children) => {
                let instruction = if removed { "delInstrText" } else { "instrText" };
                write!(
                    self.xml,
                    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:{instruction} xml:space="preserve"> REF b{target} \h </w:{instruction}></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>"#
                )
                .unwrap();
                self.inlines(children, removed);
                self.xml
                    .push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);
            }
            Inline::Comment(id, children) => {
                self.comment_ids.insert(*id);
                write!(self.xml, r#"<w:commentRangeStart w:id="{id}"/>"#).unwrap();
                self.inlines(children, removed);
                write!(
                    self.xml,
                    r#"<w:commentRangeEnd w:id="{id}"/><w:r><w:commentReference w:id="{id}"/></w:r>"#
                )
                .unwrap();
            }
            Inline::Bookmark(id, children) => {
                self.bookmark_start(*id);
                self.inlines(children, removed);
                self.bookmark_end(*id);
            }
            Inline::ProofErr => self.xml.push_str(r#"<w:proofErr w:type="spellStart"/>"#),
            Inline::Textbox(blocks) => {
                self.xml
                    .push_str("<w:r><w:pict><v:shape><v:textbox><w:txbxContent>");
                self.blocks(blocks);
                self.xml
                    .push_str("</w:txbxContent></v:textbox></v:shape></w:pict></w:r>");
            }
        }
    }

    fn paragraph(&mut self, paragraph: &Paragraph) {
        self.xml.push_str("<w:p>");
        let mark_properties = paragraph.mark.is_some() || paragraph.mark_property_change;
        if mark_properties || paragraph.property_change {
            self.xml.push_str("<w:pPr>");
            if mark_properties {
                self.xml.push_str("<w:rPr>");
                if let Some(mark) = paragraph.mark {
                    let attributes = self.revision_attributes(Author::Generated);
                    write!(self.xml, "<w:{} {attributes}/>", mark.tag()).unwrap();
                }
                if paragraph.mark_property_change {
                    self.change("rPrChange", "<w:rPr/>");
                }
                self.xml.push_str("</w:rPr>");
            }
            if paragraph.property_change {
                self.change("pPrChange", "<w:pPr/>");
            }
            self.xml.push_str("</w:pPr>");
        }
        self.inlines(&paragraph.inlines, false);
        self.xml.push_str("</w:p>");
    }

    fn table(&mut self, table: &Table) {
        self.xml
            .push_str(r#"<w:tbl><w:tblPr><w:tblStyle w:val="MissingTableStyle"/>"#);
        if table.property_changes.table {
            self.change("tblPrChange", "<w:tblPr/>");
        }
        self.xml.push_str("</w:tblPr><w:tblGrid/>");
        if let Some((TableLevel::Table, id)) = table.bookmark {
            self.bookmark_start(id);
        }
        self.xml.push_str("<w:tr>");
        let row_revision = match table.revision {
            Some(TableRevision::Row(kind)) => Some(kind),
            _ => None,
        };
        if row_revision.is_some() || table.property_changes.row {
            self.xml.push_str("<w:trPr>");
            if let Some(kind) = row_revision {
                let attributes = self.revision_attributes(Author::Generated);
                write!(self.xml, "<w:{} {attributes}/>", kind.tag()).unwrap();
            }
            if table.property_changes.row {
                self.change("trPrChange", "<w:trPr/>");
            }
            self.xml.push_str("</w:trPr>");
        }
        if let Some((TableLevel::Row, id)) = table.bookmark {
            self.bookmark_start(id);
        }
        for cell in &table.cells {
            self.xml.push_str("<w:tc>");
            let cell_revision = match table.revision {
                Some(TableRevision::CellInsertion) => Some("cellIns"),
                Some(TableRevision::CellDeletion) => Some("cellDel"),
                Some(TableRevision::CellMerge) => Some("cellMerge"),
                Some(TableRevision::Row(_)) | None => None,
            };
            if cell_revision.is_some() || table.property_changes.cell {
                self.xml.push_str("<w:tcPr>");
                if let Some(tag) = cell_revision {
                    let attributes = self.revision_attributes(Author::Generated);
                    let merge = if tag == "cellMerge" {
                        r#" w:vMerge="cont""#
                    } else {
                        ""
                    };
                    write!(self.xml, "<w:{tag} {attributes}{merge}/>").unwrap();
                }
                if table.property_changes.cell {
                    self.change("tcPrChange", "<w:tcPr/>");
                }
                self.xml.push_str("</w:tcPr>");
            }
            self.blocks(cell);
            self.xml.push_str("</w:tc>");
        }
        if let Some((TableLevel::Row, id)) = table.bookmark {
            self.bookmark_end(id);
        }
        self.xml.push_str("</w:tr>");
        if let Some((TableLevel::Table, id)) = table.bookmark {
            self.bookmark_end(id);
        }
        self.xml.push_str("</w:tbl>");
    }

    fn blocks(&mut self, blocks: &[Block]) {
        for block in blocks {
            match block {
                Block::Paragraph(paragraph) => self.paragraph(paragraph),
                Block::Table(table) => self.table(table),
                Block::Sdt(children) => {
                    self.xml.push_str("<w:sdt><w:sdtPr/><w:sdtContent>");
                    self.blocks(children);
                    self.xml.push_str("</w:sdtContent></w:sdt>");
                }
                Block::CustomXml(children) => {
                    self.xml.push_str(r#"<w:customXml w:element="section">"#);
                    self.blocks(children);
                    self.xml.push_str("</w:customXml>");
                }
                Block::Bookmark(id, children) => {
                    self.bookmark_start(*id);
                    self.blocks(children);
                    self.bookmark_end(*id);
                }
                Block::Comment(id, children) => {
                    self.comment_ids.insert(*id);
                    write!(self.xml, r#"<w:commentRangeStart w:id="{id}"/>"#).unwrap();
                    self.blocks(children);
                    write!(self.xml, r#"<w:commentRangeEnd w:id="{id}"/>"#).unwrap();
                }
            }
        }
    }
}

/// Serializes the document part and a comments part. `comment_ids` lists
/// comments a transformed document no longer anchors, so every variant of a
/// document carries the same comments part.
fn write_document(document: &Document, comment_ids: &BTreeSet<usize>) -> (String, String) {
    let (xml, comments, _) = serialize(document, comment_ids);
    (xml, comments)
}

fn comment_ids(document: &Document) -> BTreeSet<usize> {
    serialize(document, &BTreeSet::new()).2
}

fn serialize(
    document: &Document,
    comment_ids: &BTreeSet<usize>,
) -> (String, String, BTreeSet<usize>) {
    let mut markup = Markup {
        xml: format!(r#"<w:document xmlns:w="{W}" xmlns:v="{V}"><w:body>"#),
        comment_ids: comment_ids.clone(),
        next_revision: 1,
    };
    markup.blocks(&document.blocks);
    markup.xml.push_str("<w:sectPr>");
    if document.section_change {
        markup.change("sectPrChange", "<w:sectPr/>");
    }
    markup.xml.push_str("</w:sectPr></w:body></w:document>");
    let mut comments = format!(r#"<w:comments xmlns:w="{W}">"#);
    for id in &markup.comment_ids {
        write!(
            comments,
            r#"<w:comment w:id="{id}" w:author="Reviewer"><w:p><w:r><w:t>c{id}</w:t></w:r></w:p></w:comment>"#
        )
        .unwrap();
    }
    comments.push_str("</w:comments>");
    (markup.xml, comments, markup.comment_ids)
}

fn package(document: &str, comments: &str) -> Vec<u8> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (path, contents) in [
        ("word/document.xml", document),
        ("word/comments.xml", comments),
        ("word/styles.xml", STYLES),
        ("word/_rels/document.xml.rels", RELATIONSHIPS),
    ] {
        archive.start_file(path, options).unwrap();
        archive.write_all(contents.as_bytes()).unwrap();
    }
    archive.finish().unwrap().into_inner()
}

fn allocate(facts: ParagraphIdentityFacts<'_>) -> Result<InternalParagraphId, ProjectionError> {
    InternalParagraphId::new(format!("internal-{}", facts.ordinal))
}

fn project_xml(document: &str, comments: &str, view: RevisionView) -> DocumentPackageProjection {
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

fn project(
    document: &Document,
    comment_ids: &BTreeSet<usize>,
    view: RevisionView,
) -> DocumentPackageProjection {
    let (xml, comments) = write_document(document, comment_ids);
    project_xml(&xml, &comments, view)
}

// ---------------------------------------------------------------------------
// Context table, bound to the parser's element dispatch.

#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
enum Element {
    Paragraph,
    ParagraphProperties,
    Table,
    TableProperties,
    TablePropertyExceptions,
    TableStyle,
    TableGrid,
    Row,
    RowProperties,
    Cell,
    CellProperties,
    SectionProperties,
    Sdt,
    SdtContent,
    CustomXml,
    SmartTag,
    Hyperlink,
    SimpleField,
    ComplexField,
    Instruction,
    DeletedInstruction,
    Run,
    RunProperties,
    Bold,
    Text,
    DeletedText,
    Tab,
    Break,
    FootnoteReference,
    EndnoteReference,
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    RunPropertiesChange,
    ParagraphPropertiesChange,
    TablePropertiesChange,
    TablePropertiesExceptionChange,
    TableGridChange,
    TableRowPropertiesChange,
    TableCellPropertiesChange,
    SectionPropertiesChange,
    CellInsertion,
    CellDeletion,
    CellMerge,
    CommentRangeStart,
    CommentRangeEnd,
    CommentReference,
    BookmarkStart,
    BookmarkEnd,
    ProofErr,
    Textbox,
}

impl Element {
    const ALL: [Self; 52] = [
        Self::Paragraph,
        Self::ParagraphProperties,
        Self::Table,
        Self::TableProperties,
        Self::TablePropertyExceptions,
        Self::TableStyle,
        Self::TableGrid,
        Self::Row,
        Self::RowProperties,
        Self::Cell,
        Self::CellProperties,
        Self::SectionProperties,
        Self::Sdt,
        Self::SdtContent,
        Self::CustomXml,
        Self::SmartTag,
        Self::Hyperlink,
        Self::SimpleField,
        Self::ComplexField,
        Self::Instruction,
        Self::DeletedInstruction,
        Self::Run,
        Self::RunProperties,
        Self::Bold,
        Self::Text,
        Self::DeletedText,
        Self::Tab,
        Self::Break,
        Self::FootnoteReference,
        Self::EndnoteReference,
        Self::Insertion,
        Self::Deletion,
        Self::MoveFrom,
        Self::MoveTo,
        Self::RunPropertiesChange,
        Self::ParagraphPropertiesChange,
        Self::TablePropertiesChange,
        Self::TablePropertiesExceptionChange,
        Self::TableGridChange,
        Self::TableRowPropertiesChange,
        Self::TableCellPropertiesChange,
        Self::SectionPropertiesChange,
        Self::CellInsertion,
        Self::CellDeletion,
        Self::CellMerge,
        Self::CommentRangeStart,
        Self::CommentRangeEnd,
        Self::CommentReference,
        Self::BookmarkStart,
        Self::BookmarkEnd,
        Self::ProofErr,
        Self::Textbox,
    ];

    const fn local_name(self) -> &'static str {
        match self {
            Self::Paragraph => "p",
            Self::ParagraphProperties => "pPr",
            Self::Table => "tbl",
            Self::TableProperties => "tblPr",
            Self::TablePropertyExceptions => "tblPrEx",
            Self::TableStyle => "tblStyle",
            Self::TableGrid => "tblGrid",
            Self::Row => "tr",
            Self::RowProperties => "trPr",
            Self::Cell => "tc",
            Self::CellProperties => "tcPr",
            Self::SectionProperties => "sectPr",
            Self::Sdt => "sdt",
            Self::SdtContent => "sdtContent",
            Self::CustomXml => "customXml",
            Self::SmartTag => "smartTag",
            Self::Hyperlink => "hyperlink",
            Self::SimpleField => "fldSimple",
            Self::ComplexField => "fldChar",
            Self::Instruction => "instrText",
            Self::DeletedInstruction => "delInstrText",
            Self::Run => "r",
            Self::RunProperties => "rPr",
            Self::Bold => "b",
            Self::Text => "t",
            Self::DeletedText => "delText",
            Self::Tab => "tab",
            Self::Break => "br",
            Self::FootnoteReference => "footnoteReference",
            Self::EndnoteReference => "endnoteReference",
            Self::Insertion => "ins",
            Self::Deletion => "del",
            Self::MoveFrom => "moveFrom",
            Self::MoveTo => "moveTo",
            Self::RunPropertiesChange => "rPrChange",
            Self::ParagraphPropertiesChange => "pPrChange",
            Self::TablePropertiesChange => "tblPrChange",
            Self::TablePropertiesExceptionChange => "tblPrExChange",
            Self::TableGridChange => "tblGridChange",
            Self::TableRowPropertiesChange => "trPrChange",
            Self::TableCellPropertiesChange => "tcPrChange",
            Self::SectionPropertiesChange => "sectPrChange",
            Self::CellInsertion => "cellIns",
            Self::CellDeletion => "cellDel",
            Self::CellMerge => "cellMerge",
            Self::CommentRangeStart => "commentRangeStart",
            Self::CommentRangeEnd => "commentRangeEnd",
            Self::CommentReference => "commentReference",
            Self::BookmarkStart => "bookmarkStart",
            Self::BookmarkEnd => "bookmarkEnd",
            Self::ProofErr => "proofErr",
            Self::Textbox => "txbxContent",
        }
    }

    fn from_local_name(name: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|element| element.local_name() == name)
    }

    /// The containers the schema admits this element in, restricted to the
    /// placements the parser handles. Every listed context is a coverage
    /// receipt row.
    fn contexts(self) -> Vec<Context> {
        use Context as C;
        let block = [
            C::Body,
            C::TableCell,
            C::BlockSdt,
            C::BlockCustomXml,
            C::Textbox,
        ];
        let run_content = [
            C::Paragraph,
            C::Insertion,
            C::Deletion,
            C::MoveFrom,
            C::MoveTo,
            C::InlineSdt,
            C::SmartTag,
            C::InlineCustomXml,
            C::FieldResult,
            C::Hyperlink,
            C::SimpleField,
        ];
        let union = |parts: &[&[Context]]| parts.concat();
        match self {
            Self::Paragraph | Self::Table => block.to_vec(),
            Self::Sdt
            | Self::SdtContent
            | Self::CustomXml
            | Self::CommentRangeStart
            | Self::CommentRangeEnd => union(&[&block, &run_content]),
            Self::BookmarkStart | Self::BookmarkEnd => {
                union(&[&block, &run_content, &[C::Table, C::TableRow]])
            }
            Self::SmartTag | Self::Run | Self::ProofErr | Self::ComplexField => {
                run_content.to_vec()
            }
            // `CT_RunTrackChange` admits run content but not hyperlinks or
            // simple fields.
            Self::Hyperlink | Self::SimpleField => run_content
                .into_iter()
                .filter(|context| !Tracked::ALL.map(Context::revision).contains(context))
                .collect(),
            Self::Insertion | Self::Deletion => union(&[
                &run_content,
                &[C::ParagraphMarkProperties, C::TableRowProperties],
            ]),
            Self::MoveFrom | Self::MoveTo => union(&[&run_content, &[C::ParagraphMarkProperties]]),
            Self::Text
            | Self::DeletedText
            | Self::Instruction
            | Self::DeletedInstruction
            | Self::Tab
            | Self::Break
            | Self::FootnoteReference
            | Self::EndnoteReference
            | Self::CommentReference
            | Self::Textbox => vec![C::Run],
            Self::RunProperties => vec![C::Run, C::ParagraphProperties],
            Self::Bold => vec![C::RunProperties],
            Self::RunPropertiesChange => vec![C::RunProperties, C::ParagraphMarkProperties],
            Self::ParagraphProperties => vec![C::Paragraph],
            Self::ParagraphPropertiesChange => vec![C::ParagraphProperties],
            Self::TableProperties | Self::TableGrid | Self::Row => vec![C::Table],
            Self::TableStyle | Self::TablePropertiesChange => vec![C::TableProperties],
            Self::TableGridChange => vec![C::TableGrid],
            Self::RowProperties | Self::TablePropertyExceptions | Self::Cell => vec![C::TableRow],
            Self::TablePropertiesExceptionChange => vec![C::TablePropertyExceptions],
            Self::TableRowPropertiesChange => vec![C::TableRowProperties],
            Self::CellProperties => vec![C::TableCell],
            Self::TableCellPropertiesChange
            | Self::CellInsertion
            | Self::CellDeletion
            | Self::CellMerge => vec![C::TableCellProperties],
            Self::SectionProperties => vec![C::Body],
            Self::SectionPropertiesChange => vec![C::SectionProperties],
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
enum Context {
    Body,
    TableCell,
    BlockSdt,
    BlockCustomXml,
    Textbox,
    Table,
    TableRow,
    Paragraph,
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    Hyperlink,
    InlineSdt,
    SmartTag,
    InlineCustomXml,
    SimpleField,
    FieldResult,
    Run,
    RunProperties,
    ParagraphProperties,
    ParagraphMarkProperties,
    TableProperties,
    TablePropertyExceptions,
    TableRowProperties,
    TableCellProperties,
    TableGrid,
    SectionProperties,
}

impl Context {
    const fn revision(kind: Tracked) -> Self {
        match kind {
            Tracked::Insertion => Self::Insertion,
            Tracked::Deletion => Self::Deletion,
            Tracked::MoveFrom => Self::MoveFrom,
            Tracked::MoveTo => Self::MoveTo,
        }
    }

    const fn is_block(self) -> bool {
        matches!(
            self,
            Self::Body | Self::TableCell | Self::BlockSdt | Self::BlockCustomXml | Self::Textbox
        )
    }
}

impl Tracked {
    const ALL: [Self; 4] = [
        Self::Insertion,
        Self::Deletion,
        Self::MoveFrom,
        Self::MoveTo,
    ];
}

/// Dispatched names not generated in this suite.
const UNGENERATED_DISPATCH: &[&str] = &[
    "body",
    "pStyle",
    "ind",
    "jc",
    "outlineLvl",
    "numPr",
    "numId",
    "ilvl",
    "rStyle",
    "bCs",
    "cs",
    "rtl",
    "highlight",
    "vertAlign",
    "vanish",
    "showingPlcHdr",
    "ptab",
    "cr",
    "softHyphen",
    "noBreakHyphen",
    "sym",
    "oMath",
    "oMathPara",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
];

/// Generated elements the parser handles through its default arm: they are
/// transparent containers or markers whose content the walk still visits.
const DEFAULT_ARM: &[Element] = &[Element::SmartTag, Element::ProofErr];

/// Element names the parser dispatches on: byte-string patterns in
/// `ProjectionState::start` and in the revision classifiers it calls.
/// Attribute names (`attribute(reader, element, b"...")`) are excluded.
fn parser_dispatch() -> BTreeSet<String> {
    const SOURCE: &str = include_str!("../src/projection/ooxml.rs");
    let find_section = |start: &str, end: &str| {
        let from = SOURCE
            .find(start)
            .unwrap_or_else(|| panic!("{start} moved"));
        let to = from + SOURCE[from..].find(end).unwrap();
        &SOURCE[from..to]
    };
    let sections = [
        find_section("    fn start(", "    fn end("),
        find_section("const fn revision_fact_kind(", "\n}\n"),
        find_section("fn is_change_snapshot(", "\n}\n"),
        find_section("fn is_unsupported_revision_markup(", "\n}\n"),
    ];
    let mut names = BTreeSet::new();
    for section in sections {
        let mut rest = section;
        while let Some(index) = rest.find("b\"") {
            let preceding = &rest[..index];
            rest = &rest[index + 2..];
            let end = rest.find('"').unwrap();
            let name = &rest[..end];
            rest = &rest[end + 1..];
            if name.is_empty() || preceding.ends_with("element, ") {
                continue;
            }
            names.insert(name.to_owned());
        }
    }
    names
}

#[test]
fn context_table_matches_the_parser_dispatch() {
    let dispatched = parser_dispatch();
    assert!(
        dispatched.contains("ins") && dispatched.contains("rPrChange"),
        "the dispatch scan found the parser's element patterns: {dispatched:?}"
    );
    let ungenerated = UNGENERATED_DISPATCH
        .iter()
        .map(|name| (*name).to_owned())
        .collect::<BTreeSet<_>>();
    let generated = Element::ALL
        .iter()
        .filter(|element| !DEFAULT_ARM.contains(element))
        .map(|element| element.local_name().to_owned())
        .collect::<BTreeSet<_>>();
    let declared = generated
        .union(&ungenerated)
        .cloned()
        .collect::<BTreeSet<_>>();
    assert_eq!(
        dispatched.difference(&declared).collect::<Vec<_>>(),
        Vec::<&String>::new(),
        "every dispatched element needs an Element row or a neutral UNGENERATED_DISPATCH entry"
    );
    assert_eq!(
        declared.difference(&dispatched).collect::<Vec<_>>(),
        Vec::<&String>::new(),
        "an element the parser no longer dispatches must move to DEFAULT_ARM or go"
    );
    for element in DEFAULT_ARM {
        assert!(
            !dispatched.contains(element.local_name()),
            "{element:?} is dispatched now; remove it from DEFAULT_ARM"
        );
    }
    assert!(
        generated.is_disjoint(&ungenerated),
        "an element is either generated or declared ungenerated"
    );
    let distinct = Element::ALL
        .iter()
        .map(|element| element.local_name())
        .collect::<BTreeSet<_>>();
    assert_eq!(
        distinct.len(),
        Element::ALL.len(),
        "Element::ALL is distinct"
    );
}

// ---------------------------------------------------------------------------
// Coverage receipt: an independent XML walk records which (element, context)
// rows a document actually contains.

struct ObservedFrame {
    container: Option<Context>,
    open_field_results: usize,
    name: String,
}

fn local_name(element: &BytesStart<'_>) -> String {
    String::from_utf8(element.local_name().as_ref().to_vec()).unwrap()
}

fn attribute_value(element: &BytesStart<'_>, name: &str) -> Option<String> {
    element.attributes().find_map(|attribute| {
        let attribute = attribute.unwrap();
        (attribute.key.local_name().as_ref() == name.as_bytes())
            .then(|| String::from_utf8(attribute.value.to_vec()).unwrap())
    })
}

fn current_context(stack: &[ObservedFrame], skip_runs: bool) -> Option<Context> {
    stack.iter().rev().find_map(|frame| {
        let container = frame.container?;
        if skip_runs && container == Context::Run {
            return None;
        }
        Some(if frame.open_field_results > 0 {
            Context::FieldResult
        } else {
            container
        })
    })
}

fn is_property_snapshot(name: &str) -> bool {
    name.ends_with("PrChange") || name == "tblGridChange" || name == "tblPrExChange"
}

fn observe(xml: &str, rows: &mut BTreeSet<(Element, Context)>) {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut stack: Vec<ObservedFrame> = Vec::new();
    let mut snapshot_depth = 0_usize;
    loop {
        let (element, empty) = match reader.read_event().unwrap() {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                let frame = stack.pop().unwrap();
                if is_property_snapshot(&frame.name) {
                    snapshot_depth -= 1;
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        let name = local_name(&element);
        let parent_is_paragraph_properties = stack.last().is_some_and(|frame| frame.name == "pPr");
        let field_type = attribute_value(&element, "fldCharType");
        let context = current_context(&stack, name == "fldChar");
        let recorded = name != "fldChar" || field_type.as_deref() == Some("begin");
        if snapshot_depth == 0
            && recorded
            && let (Some(row_element), Some(context)) = (Element::from_local_name(&name), context)
        {
            rows.insert((row_element, context));
        }
        if name == "fldChar" {
            let container = stack
                .iter_mut()
                .rev()
                .find(|frame| frame.container.is_some_and(|c| c != Context::Run))
                .unwrap();
            match field_type.as_deref() {
                Some("separate") => container.open_field_results += 1,
                Some("end") => container.open_field_results -= 1,
                _ => {}
            }
        }
        let container = match name.as_str() {
            "body" => Some(Context::Body),
            "tbl" => Some(Context::Table),
            "tr" => Some(Context::TableRow),
            "tc" => Some(Context::TableCell),
            "p" => Some(Context::Paragraph),
            "r" => Some(Context::Run),
            "hyperlink" => Some(Context::Hyperlink),
            "smartTag" => Some(Context::SmartTag),
            "fldSimple" => Some(Context::SimpleField),
            "txbxContent" => Some(Context::Textbox),
            "pPr" => Some(Context::ParagraphProperties),
            "trPr" => Some(Context::TableRowProperties),
            "tcPr" => Some(Context::TableCellProperties),
            "tblPr" => Some(Context::TableProperties),
            "tblPrEx" => Some(Context::TablePropertyExceptions),
            "tblGrid" => Some(Context::TableGrid),
            "sectPr" => Some(Context::SectionProperties),
            "rPr" => Some(if parent_is_paragraph_properties {
                Context::ParagraphMarkProperties
            } else {
                Context::RunProperties
            }),
            "ins" => Some(Context::Insertion),
            "del" => Some(Context::Deletion),
            "moveFrom" => Some(Context::MoveFrom),
            "moveTo" => Some(Context::MoveTo),
            // An `w:sdtContent` takes the context of its `w:sdt`.
            "sdtContent" | "customXml" => {
                let block = context.is_some_and(Context::is_block);
                Some(match (name.as_str(), block) {
                    ("sdtContent", true) => Context::BlockSdt,
                    ("sdtContent", false) => Context::InlineSdt,
                    (_, true) => Context::BlockCustomXml,
                    _ => Context::InlineCustomXml,
                })
            }
            _ => None,
        };
        if is_property_snapshot(&name) {
            snapshot_depth += 1;
        }
        if empty {
            if is_property_snapshot(&name) {
                snapshot_depth -= 1;
            }
            continue;
        }
        stack.push(ObservedFrame {
            container,
            open_field_results: 0,
            name,
        });
    }
}

/// Placements not generated in this suite.
const EXCLUDED_ROWS: &[(Element, Context)] = &[
    (Element::Paragraph, Context::Textbox),
    (Element::Table, Context::TableCell),
    (Element::Table, Context::Textbox),
    (Element::Sdt, Context::TableCell),
    (Element::Sdt, Context::Textbox),
    (Element::SdtContent, Context::TableCell),
    (Element::SdtContent, Context::Textbox),
    (Element::CustomXml, Context::TableCell),
    (Element::CustomXml, Context::Textbox),
    (Element::CommentRangeStart, Context::TableCell),
    (Element::CommentRangeStart, Context::Textbox),
    (Element::CommentRangeEnd, Context::TableCell),
    (Element::CommentRangeEnd, Context::Textbox),
    (Element::BookmarkStart, Context::TableCell),
    (Element::BookmarkStart, Context::Textbox),
    (Element::BookmarkStart, Context::Table),
    (Element::BookmarkStart, Context::TableRow),
    (Element::BookmarkEnd, Context::TableCell),
    (Element::BookmarkEnd, Context::Textbox),
    (Element::BookmarkEnd, Context::Table),
    (Element::BookmarkEnd, Context::TableRow),
    (Element::Insertion, Context::TableRowProperties),
    (Element::Deletion, Context::TableRowProperties),
    (Element::Textbox, Context::Run),
    (Element::CellInsertion, Context::TableCellProperties),
    (Element::CellDeletion, Context::TableCellProperties),
    (Element::CellMerge, Context::TableCellProperties),
    (Element::SectionPropertiesChange, Context::SectionProperties),
];

#[test]
fn snapshot_receipt_resumes_after_table_property_exceptions() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tblPrEx><w:tblPrExChange w:id="1" w:author="A"><w:tblPrEx/></w:tblPrExChange></w:tblPrEx><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#
    );
    let mut rows = BTreeSet::new();
    observe(&xml, &mut rows);
    assert!(rows.contains(&(
        Element::TablePropertiesExceptionChange,
        Context::TablePropertyExceptions
    )));
    assert!(rows.contains(&(Element::Cell, Context::TableRow)));
    assert!(rows.contains(&(Element::Bold, Context::RunProperties)));
    assert!(rows.contains(&(Element::Text, Context::Run)));
    assert!(!rows.contains(&(
        Element::TablePropertyExceptions,
        Context::TablePropertyExceptions
    )));
}

fn expected_rows() -> BTreeSet<(Element, Context)> {
    Element::ALL
        .iter()
        .flat_map(|element| {
            element
                .contexts()
                .into_iter()
                .map(move |context| (*element, context))
        })
        .filter(|row| !EXCLUDED_ROWS.contains(row))
        .collect()
}

/// Generated and excluded placements partition the table; the receipt
/// exercises every generated placement and none of the excluded placements.
#[test]
fn generator_reaches_every_context_row() {
    let mut runner = TestRunner::deterministic();
    let strategy = prop_oneof![document().boxed(), property_document().boxed()];
    let mut observed = BTreeSet::new();
    for _ in 0..RECEIPT_CASES {
        let document = strategy.new_tree(&mut runner).unwrap().current();
        let (xml, _) = write_document(&document, &BTreeSet::new());
        observe(&xml, &mut observed);
    }
    let table_strategy = table_property_document();
    for _ in 0..RECEIPT_CASES {
        let table = table_strategy.new_tree(&mut runner).unwrap().current();
        observe(&table.xml(), &mut observed);
    }
    let expected = expected_rows();
    let excluded = EXCLUDED_ROWS.iter().copied().collect::<BTreeSet<_>>();
    assert_eq!(
        excluded.len(),
        EXCLUDED_ROWS.len(),
        "excluded rows are distinct"
    );
    let declared = Element::ALL
        .iter()
        .flat_map(|element| {
            element
                .contexts()
                .into_iter()
                .map(move |context| (*element, context))
        })
        .collect::<BTreeSet<_>>();
    assert!(
        excluded.is_subset(&declared),
        "excluded rows belong to the context table"
    );
    assert!(
        observed.is_disjoint(&excluded),
        "the generator emits no excluded placements"
    );
    assert_eq!(
        expected.union(&excluded).copied().collect::<BTreeSet<_>>(),
        declared
    );
    let missing = expected.difference(&observed).collect::<Vec<_>>();
    let unexpected = observed.difference(&expected).collect::<Vec<_>>();
    assert!(
        missing.is_empty() && unexpected.is_empty(),
        "coverage receipt: {} of {} rows; missing {missing:?}; outside the table {unexpected:?}",
        expected.len() - missing.len(),
        expected.len(),
    );
}

// ---------------------------------------------------------------------------
// Revision locations in the independent model.

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Site {
    Inline,
    PropertyChange,
    TableStructure,
    Section,
    Textbox,
}

// ---------------------------------------------------------------------------
// Reference model of the projection.

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct SourcePoint {
    paragraph: usize,
    utf8: u32,
    utf16: u32,
}

#[derive(Debug)]
struct ExpectedRevision {
    kind: RevisionFactKind,
    site: Site,
    located: Option<(SourcePoint, SourcePoint, RevisionPayload)>,
}

struct SourceParagraph {
    text: String,
    utf16: u32,
    formatting: Vec<TextFormattingSpan>,
    cell: Option<(usize, usize)>,
    mark: Option<Tracked>,
}

struct Model {
    view: RevisionView,
    paragraphs: Vec<SourceParagraph>,
    revisions: Vec<ExpectedRevision>,
    open_revisions: Vec<(usize, String)>,
    comments: BTreeMap<usize, Option<(SourcePoint, SourcePoint)>>,
    next_table: usize,
}

/// The model's expectation for one document in one view.
struct Expected {
    texts: Vec<String>,
    formatting: Vec<Vec<TextFormattingSpan>>,
    origins: Vec<(usize, u32, u32)>,
    revisions: Vec<ExpectedRevision>,
    comments: BTreeMap<usize, Option<(SourcePoint, SourcePoint)>>,
}

impl Model {
    fn build(document: &Document, view: RevisionView) -> Expected {
        let mut model = Self {
            view,
            paragraphs: Vec::new(),
            revisions: Vec::new(),
            open_revisions: Vec::new(),
            comments: BTreeMap::new(),
            next_table: 0,
        };
        model.blocks(&document.blocks, None);
        if document.section_change {
            model.snapshot(RevisionFactKind::SectionPropertiesChange, Site::Section);
        }
        let (texts, origins) = model.joined();
        let mut formatting: Vec<Vec<TextFormattingSpan>> = vec![Vec::new(); texts.len()];
        for (source, &(destination, _, offset)) in model.paragraphs.iter().zip(&origins) {
            for span in &source.formatting {
                let shifted = TextFormattingSpan {
                    start_utf16: span.start_utf16 + offset,
                    end_utf16: span.end_utf16 + offset,
                    style: span.style,
                };
                let spans = &mut formatting[destination];
                if let Some(previous) = spans.last_mut()
                    && previous.end_utf16 == shifted.start_utf16
                    && previous.style == shifted.style
                {
                    previous.end_utf16 = shifted.end_utf16;
                } else {
                    spans.push(shifted);
                }
            }
        }
        Expected {
            texts,
            formatting,
            origins,
            revisions: model.revisions,
            comments: model.comments,
        }
    }

    fn here(&self) -> SourcePoint {
        let paragraph = self.paragraphs.last().unwrap();
        SourcePoint {
            paragraph: self.paragraphs.len() - 1,
            utf8: u32::try_from(paragraph.text.len()).unwrap(),
            utf16: paragraph.utf16,
        }
    }

    fn site(&mut self, kind: RevisionFactKind, site: Site) {
        self.revisions.push(ExpectedRevision {
            kind,
            site,
            located: None,
        });
    }

    fn snapshot(&mut self, kind: RevisionFactKind, site: Site) {
        self.site(kind, site);
    }

    fn push_text(&mut self, text: &str, hidden: bool) {
        for (_, covered) in &mut self.open_revisions {
            covered.push_str(text);
        }
        if hidden {
            return;
        }
        let paragraph = self.paragraphs.last_mut().unwrap();
        paragraph.text.push_str(text);
        paragraph.utf16 += u32::try_from(text.encode_utf16().count()).unwrap();
    }

    fn blocks(&mut self, blocks: &[Block], cell: Option<(usize, usize)>) {
        for block in blocks {
            match block {
                Block::Paragraph(paragraph) => self.paragraph(paragraph, cell),
                Block::Table(table) => self.table(table),
                Block::Sdt(children)
                | Block::CustomXml(children)
                | Block::Bookmark(_, children) => {
                    self.blocks(children, cell);
                }
                Block::Comment(id, children) => {
                    let start = SourcePoint {
                        paragraph: self.paragraphs.len(),
                        utf8: 0,
                        utf16: 0,
                    };
                    self.blocks(children, cell);
                    self.comments.insert(*id, Some((start, self.here())));
                }
            }
        }
    }

    fn table(&mut self, table: &Table) {
        let ordinal = self.next_table;
        self.next_table += 1;
        if table.property_changes.table {
            self.snapshot(
                RevisionFactKind::TablePropertiesChange,
                Site::TableStructure,
            );
        }
        if let Some(TableRevision::Row(kind)) = table.revision {
            self.site(kind.kind(), Site::TableStructure);
        }
        if table.property_changes.row {
            self.snapshot(
                RevisionFactKind::TableRowPropertiesChange,
                Site::TableStructure,
            );
        }
        for (column, cell) in table.cells.iter().enumerate() {
            let cell_kind = match table.revision {
                Some(TableRevision::CellInsertion) => Some(RevisionFactKind::CellInsertion),
                Some(TableRevision::CellDeletion) => Some(RevisionFactKind::CellDeletion),
                Some(TableRevision::CellMerge) => Some(RevisionFactKind::CellMerge),
                Some(TableRevision::Row(_)) | None => None,
            };
            if let Some(kind) = cell_kind {
                self.site(kind, Site::TableStructure);
            }
            if table.property_changes.cell {
                self.snapshot(
                    RevisionFactKind::TableCellPropertiesChange,
                    Site::TableStructure,
                );
            }
            self.blocks(cell, Some((ordinal, column)));
        }
    }

    fn paragraph(&mut self, paragraph: &Paragraph, cell: Option<(usize, usize)>) {
        self.paragraphs.push(SourceParagraph {
            text: String::new(),
            utf16: 0,
            formatting: Vec::new(),
            cell,
            mark: if paragraph.mark_property_change && self.view == RevisionView::Original {
                None
            } else {
                paragraph.mark
            },
        });
        let mark_index = paragraph.mark.map(|mark| {
            self.site(mark.kind(), Site::Inline);
            self.revisions.len() - 1
        });
        let start = self.here();
        let mark_snapshot = paragraph.mark_property_change.then(|| {
            self.snapshot(RevisionFactKind::RunPropertiesChange, Site::PropertyChange);
            self.revisions.len() - 1
        });
        let paragraph_snapshot = paragraph.property_change.then(|| {
            self.snapshot(
                RevisionFactKind::ParagraphPropertiesChange,
                Site::PropertyChange,
            );
            self.revisions.len() - 1
        });
        for inline in &paragraph.inlines {
            self.inline(inline, false);
        }
        let end = self.here();
        if let Some(index) = mark_snapshot {
            self.revisions[index].located = Some((end, end, RevisionPayload::FormattingOnly));
        }
        if let Some(index) = paragraph_snapshot {
            self.revisions[index].located = Some((start, end, RevisionPayload::FormattingOnly));
        }
        if let Some(index) = mark_index {
            let mark = self.here();
            self.revisions[index].located = Some((mark, mark, RevisionPayload::ParagraphMark));
        }
    }

    fn run(&mut self, run: Run, hidden: bool) {
        let start = self.here();
        let snapshot = run.property_change.then(|| {
            self.snapshot(RevisionFactKind::RunPropertiesChange, Site::PropertyChange);
            self.revisions.len() - 1
        });
        let text = match run.content {
            RunContent::Text(text) => text,
            RunContent::Tab => "\t",
            RunContent::Break => "\u{000b}",
            RunContent::FootnoteReference | RunContent::EndnoteReference => "\u{0002}",
        };
        let start_utf16 = self.paragraphs.last().unwrap().utf16;
        self.push_text(text, hidden);
        let bold = if run.property_change && self.view == RevisionView::Original {
            run.prior_bold
        } else {
            run.bold
        };
        if !hidden && bold && !text.is_empty() {
            let paragraph = self.paragraphs.last_mut().unwrap();
            let end_utf16 = paragraph.utf16;
            if let Some(last) = paragraph.formatting.last_mut()
                && last.end_utf16 == start_utf16
            {
                last.end_utf16 = end_utf16;
            } else {
                paragraph.formatting.push(TextFormattingSpan {
                    start_utf16,
                    end_utf16,
                    style: TextStyle::Bold,
                });
            }
        }
        if let Some(index) = snapshot {
            self.revisions[index].located =
                Some((start, self.here(), RevisionPayload::FormattingOnly));
        }
    }

    fn inline(&mut self, inline: &Inline, hidden: bool) {
        match inline {
            Inline::Run(run) => self.run(*run, hidden),
            Inline::Revision(revision) => {
                self.site(revision.kind.kind(), Site::Inline);
                let index = self.revisions.len() - 1;
                let start = self.here();
                self.open_revisions.push((index, String::new()));
                let hidden = hidden || revision.kind.hidden_in(self.view);
                for child in &revision.children {
                    self.inline(child, hidden);
                }
                let (_, covered) = self.open_revisions.pop().unwrap();
                let payload = if covered.is_empty() {
                    RevisionPayload::FormattingOnly
                } else {
                    RevisionPayload::Text(covered)
                };
                self.revisions[index].located = Some((start, self.here(), payload));
            }
            Inline::Hyperlink(_, children)
            | Inline::SimpleField(_, children)
            | Inline::ComplexField(_, children)
            | Inline::Sdt(children)
            | Inline::SmartTag(children)
            | Inline::CustomXml(children)
            | Inline::Bookmark(_, children) => {
                for child in children {
                    self.inline(child, hidden);
                }
            }
            Inline::Comment(id, children) => {
                let start = self.here();
                for child in children {
                    self.inline(child, hidden);
                }
                self.comments.insert(*id, Some((start, self.here())));
            }
            Inline::ProofErr => {}
            Inline::Textbox(blocks) => self.textbox_blocks(blocks),
        }
    }

    /// Textbox stories are not projected: revisions are recorded without a
    /// location and comments without an anchor.
    fn textbox_blocks(&mut self, blocks: &[Block]) {
        for block in blocks {
            match block {
                Block::Paragraph(paragraph) => {
                    if let Some(mark) = paragraph.mark {
                        self.site(mark.kind(), Site::Textbox);
                    }
                    if paragraph.mark_property_change {
                        self.site(RevisionFactKind::RunPropertiesChange, Site::Textbox);
                    }
                    if paragraph.property_change {
                        self.site(RevisionFactKind::ParagraphPropertiesChange, Site::Textbox);
                    }
                    self.textbox_inlines(&paragraph.inlines);
                }
                Block::Table(table) => {
                    if table.property_changes.table {
                        self.site(RevisionFactKind::TablePropertiesChange, Site::Textbox);
                    }
                    if let Some(TableRevision::Row(kind)) = table.revision {
                        self.site(kind.kind(), Site::Textbox);
                    }
                    if table.property_changes.row {
                        self.site(RevisionFactKind::TableRowPropertiesChange, Site::Textbox);
                    }
                    for cell in &table.cells {
                        match table.revision {
                            Some(TableRevision::CellInsertion) => {
                                self.site(RevisionFactKind::CellInsertion, Site::Textbox);
                            }
                            Some(TableRevision::CellDeletion) => {
                                self.site(RevisionFactKind::CellDeletion, Site::Textbox);
                            }
                            Some(TableRevision::CellMerge) => {
                                self.site(RevisionFactKind::CellMerge, Site::Textbox);
                            }
                            Some(TableRevision::Row(_)) | None => {}
                        }
                        if table.property_changes.cell {
                            self.site(RevisionFactKind::TableCellPropertiesChange, Site::Textbox);
                        }
                        self.textbox_blocks(cell);
                    }
                }
                Block::Sdt(children)
                | Block::CustomXml(children)
                | Block::Bookmark(_, children) => {
                    self.textbox_blocks(children);
                }
                Block::Comment(id, children) => {
                    self.comments.insert(*id, None);
                    self.textbox_blocks(children);
                }
            }
        }
    }

    fn textbox_inlines(&mut self, inlines: &[Inline]) {
        for inline in inlines {
            match inline {
                Inline::Run(run) => {
                    if run.property_change {
                        self.site(RevisionFactKind::RunPropertiesChange, Site::Textbox);
                    }
                }
                Inline::Revision(revision) => {
                    self.site(revision.kind.kind(), Site::Textbox);
                    self.textbox_inlines(&revision.children);
                }
                Inline::Comment(id, children) => {
                    self.comments.insert(*id, None);
                    self.textbox_inlines(children);
                }
                Inline::Hyperlink(_, children)
                | Inline::Sdt(children)
                | Inline::SmartTag(children)
                | Inline::CustomXml(children)
                | Inline::SimpleField(_, children)
                | Inline::ComplexField(_, children)
                | Inline::Bookmark(_, children) => self.textbox_inlines(children),
                Inline::ProofErr | Inline::Textbox(_) => {}
            }
        }
    }

    /// Joins paragraphs whose break the view removes, as long as both share
    /// a table cell (or both sit outside tables).
    fn joined(&self) -> (Vec<String>, Vec<(usize, u32, u32)>) {
        let mut texts: Vec<(String, u32)> = Vec::new();
        let mut origins = Vec::new();
        for index in 0..self.paragraphs.len() {
            let removed = index > 0
                && self.paragraphs[index - 1]
                    .mark
                    .is_some_and(|mark| mark.removes_paragraph_break_in(self.view));
            let compatible =
                removed && self.paragraphs[index - 1].cell == self.paragraphs[index].cell;
            let paragraph = &self.paragraphs[index];
            if compatible {
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
        (texts.into_iter().map(|(text, _)| text).collect(), origins)
    }
}

impl Expected {
    fn relocate(&self, at: SourcePoint) -> ReviewPoint {
        let (paragraph_ordinal, utf8, utf16) = self.origins[at.paragraph];
        ReviewPoint {
            paragraph_ordinal,
            utf8: utf8 + at.utf8,
            utf16: utf16 + at.utf16,
        }
    }

    fn span(&self, (start, end): (SourcePoint, SourcePoint)) -> ReviewSpan {
        ReviewSpan {
            start: self.relocate(start),
            end: self.relocate(end),
        }
    }

    fn referenced_text(&self, span: ReviewSpan) -> String {
        let text = |ordinal: usize, from: u32, to: Option<u32>| {
            let text = &self.texts[ordinal];
            let to = to.map_or(text.len(), |to| usize::try_from(to).unwrap());
            text[usize::try_from(from).unwrap()..to].to_owned()
        };
        if span.start.paragraph_ordinal == span.end.paragraph_ordinal {
            return text(
                span.start.paragraph_ordinal,
                span.start.utf8,
                Some(span.end.utf8),
            );
        }
        let mut output = text(span.start.paragraph_ordinal, span.start.utf8, None);
        for ordinal in span.start.paragraph_ordinal + 1..span.end.paragraph_ordinal {
            output.push('\n');
            output.push_str(&self.texts[ordinal]);
        }
        output.push('\n');
        output.push_str(&text(span.end.paragraph_ordinal, 0, Some(span.end.utf8)));
        output
    }
}

// ---------------------------------------------------------------------------
// Unknown-rate invariant.

fn known_revisions(projection: &DocumentPackageProjection) -> &[AttributedRevision] {
    let ReviewFactSet::Known(revisions) = &projection.review_facts.revisions else {
        panic!(
            "revision facts are unknown: {:?}",
            projection.review_facts.revisions
        );
    };
    revisions
}

/// Every generated fact family stays known and matches the model.
#[allow(clippy::too_many_lines)] // One pass over every fact family keeps the oracle auditable.
fn check_projection(document: &Document, view: RevisionView) -> Result<(), String> {
    let ids = comment_ids(document);
    let projection = project(document, &ids, view);
    check_projected_facts(document, view, &projection)
}

#[allow(clippy::too_many_lines)] // One pass over every fact family keeps the oracle auditable.
fn check_projected_facts(
    document: &Document,
    view: RevisionView,
    projection: &DocumentPackageProjection,
) -> Result<(), String> {
    let mut placements = document.clone();
    {
        map_paragraphs(&mut placements.blocks, &mut |paragraph| {
            paragraph.property_change = false;
            paragraph.mark_property_change = false;
        });
        visit_inline_lists(&mut placements.blocks, &mut |inlines| {
            for inline in inlines {
                if let Inline::Run(run) = inline {
                    run.property_change = false;
                }
            }
        });
    }
    require_scoped_placements(&placements)?;
    let expected = Model::build(document, view);
    let fail = |message: String| -> Result<(), String> { Err(format!("{view:?}: {message}")) };

    let texts = projection
        .document
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.text.clone())
        .collect::<Vec<_>>();
    if texts != expected.texts {
        return fail(format!("texts {texts:?}, expected {:?}", expected.texts));
    }
    let formatting = projection
        .document
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.formatting.clone())
        .collect::<Vec<_>>();
    if formatting != expected.formatting {
        return fail(format!(
            "formatting {formatting:?}, expected {:?}",
            expected.formatting
        ));
    }
    require_known_families(projection).map_err(|message| format!("{view:?}: {message}"))?;
    if projection.document.structural_facts.indentation != StructuralFactSet::Known(Vec::new())
        || projection.document.structural_facts.numbering_hierarchy
            != StructuralFactSet::Known(Vec::new())
        || projection.document.structural_facts.outline_levels
            != StructuralFactSet::Known(Vec::new())
    {
        return fail(
            "generated paragraphs have no indentation, numbering, or outline properties".to_owned(),
        );
    }

    let revisions = known_revisions(projection);
    let kinds = revisions
        .iter()
        .map(|revision| revision.kind)
        .collect::<Vec<_>>();
    let expected_kinds = expected
        .revisions
        .iter()
        .map(|revision| revision.kind)
        .collect::<Vec<_>>();
    if kinds != expected_kinds {
        return fail(format!(
            "revision kinds {kinds:?}, expected {expected_kinds:?}"
        ));
    }
    for (index, (revision, expectation)) in revisions.iter().zip(&expected.revisions).enumerate() {
        match (&revision.content, &expectation.located) {
            (ReviewDetail::Known(content), Some((start, end, payload))) => {
                let span = expected.span((*start, *end));
                if content.span != span || content.payload != *payload {
                    return fail(format!(
                        "revision {index} ({:?}) is {content:?}, expected {span:?} {payload:?}",
                        revision.kind
                    ));
                }
            }
            (ReviewDetail::Known(_), None) => {
                return fail(format!("revision {index} has no modeled location"));
            }
            (ReviewDetail::Unknown(reason), _) => {
                return fail(format!(
                    "revision {index} at {:?} is unknown: {reason:?}",
                    expectation.site
                ));
            }
        }
    }

    let ReviewFactSet::Known(comments) = &projection.review_facts.comments else {
        return fail("comment facts are unknown".to_owned());
    };
    let mut emitted_ids = comments
        .iter()
        .map(|comment| comment.comment_id.parse::<usize>().unwrap())
        .collect::<Vec<_>>();
    emitted_ids.sort_unstable();
    let expected_ids = expected.comments.keys().copied().collect::<Vec<_>>();
    if emitted_ids != expected_ids {
        return fail(format!(
            "comment ids {emitted_ids:?}, expected {expected_ids:?}"
        ));
    }
    for comment in comments {
        let id = comment.comment_id.parse::<usize>().unwrap();
        let Some(anchor) = expected.comments.get(&id) else {
            return fail(format!("comment {id} is not in the document"));
        };
        match (anchor, &comment.content) {
            (Some(anchor), ReviewDetail::Known(content)) => {
                let anchor = expected.span(*anchor);
                let expected_content = CommentContent {
                    anchor,
                    comment_text: format!("c{id}"),
                    referenced_text: expected.referenced_text(anchor),
                };
                if *content != expected_content {
                    return fail(format!(
                        "comment {id} is {content:?}, expected {expected_content:?}"
                    ));
                }
            }
            (Some(_), ReviewDetail::Unknown(reason)) => {
                return fail(format!("comment {id} is unknown: {reason:?}"));
            }
            (None, _) => return fail(format!("comment {id} has no modeled anchor")),
        }
    }
    Ok(())
}

proptest! {
    #![proptest_config(config(96))]

    #[test]
    fn generated_property_documents_match_the_model(document in property_document()) {
        for view in VIEWS {
            let checked = check_projection(&document, view);
            prop_assert!(checked.is_ok(), "{}", checked.unwrap_err());
        }
    }

    /// Every generated fact stays known and matches the model.
    #[test]
    fn generated_documents_match_the_model(document in document()) {
        for view in VIEWS {
            let checked = check_projection(&document, view);
            prop_assert!(checked.is_ok(), "{}", checked.unwrap_err());
        }
    }
}

proptest! {
    #![proptest_config(config(64))]

    /// A projection cannot drop or duplicate generated comments unnoticed.
    #[test]
    fn comment_oracle_rejects_missing_and_duplicate_facts(mut document in document()) {
        document.blocks.push(Block::Comment(0, vec![Block::Paragraph(empty_paragraph())]));
        sanitize_document(&mut document);
        let ids = comment_ids(&document);
        for view in VIEWS {
            let mut projection = project(&document, &ids, view);
            prop_assert!(check_projected_facts(&document, view, &projection).is_ok());
            let ReviewFactSet::Known(comments) = &mut projection.review_facts.comments else {
                panic!("generated comment facts are known");
            };
            let comment = comments.pop().expect("the generator includes a comment");
            prop_assert!(check_projected_facts(&document, view, &projection).is_err());
            let ReviewFactSet::Known(restored_comments) = &mut projection.review_facts.comments else {
                panic!("comment facts stay known while mutating their entries");
            };
            restored_comments.push(comment.clone());
            restored_comments.push(comment);
            prop_assert!(check_projected_facts(&document, view, &projection).is_err());
        }
    }
}

// ---------------------------------------------------------------------------
// Metamorphic relations.

/// Splits every splittable text run into two runs with identical formatting.
fn split_runs(inlines: &mut Vec<Inline>) {
    let mut output = Vec::with_capacity(inlines.len());
    for mut inline in std::mem::take(inlines) {
        match &mut inline {
            Inline::Run(run) if !run.property_change => {
                let halves = match run.content {
                    RunContent::Text("bc") => Some(("b", "c")),
                    RunContent::Text("x y") => Some(("x", " y")),
                    _ => None,
                };
                if let Some((first, second)) = halves {
                    output.push(Inline::Run(Run {
                        content: RunContent::Text(first),
                        ..*run
                    }));
                    output.push(Inline::Run(Run {
                        content: RunContent::Text(second),
                        ..*run
                    }));
                    continue;
                }
            }
            Inline::Run(_) | Inline::ProofErr => {}
            Inline::Revision(revision) => split_runs(&mut revision.children),
            Inline::Hyperlink(_, children)
            | Inline::Sdt(children)
            | Inline::SmartTag(children)
            | Inline::CustomXml(children)
            | Inline::SimpleField(_, children)
            | Inline::ComplexField(_, children)
            | Inline::Comment(_, children)
            | Inline::Bookmark(_, children) => split_runs(children),
            Inline::Textbox(blocks) => map_paragraphs(blocks, &mut |paragraph| {
                split_runs(&mut paragraph.inlines);
            }),
        }
        output.push(inline);
    }
    *inlines = output;
}

fn map_paragraphs(blocks: &mut [Block], apply: &mut impl FnMut(&mut Paragraph)) {
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => apply(paragraph),
            Block::Table(table) => {
                for cell in &mut table.cells {
                    map_paragraphs(cell, apply);
                }
            }
            Block::Sdt(children)
            | Block::CustomXml(children)
            | Block::Bookmark(_, children)
            | Block::Comment(_, children) => map_paragraphs(children, apply),
        }
    }
}

/// Visits inline lists outside revision ancestry.
fn visit_inline_lists(blocks: &mut [Block], apply: &mut impl FnMut(&mut Vec<Inline>)) {
    fn inlines(
        list: &mut Vec<Inline>,
        inside_revision: bool,
        apply: &mut impl FnMut(&mut Vec<Inline>),
    ) {
        if !inside_revision {
            apply(list);
        }
        for inline in list {
            match inline {
                Inline::Revision(revision) => inlines(&mut revision.children, true, apply),
                Inline::Hyperlink(_, children)
                | Inline::Sdt(children)
                | Inline::SmartTag(children)
                | Inline::CustomXml(children)
                | Inline::SimpleField(_, children)
                | Inline::ComplexField(_, children)
                | Inline::Comment(_, children)
                | Inline::Bookmark(_, children) => inlines(children, inside_revision, apply),
                Inline::Run(_) | Inline::ProofErr | Inline::Textbox(_) => {}
            }
        }
    }
    map_paragraphs(blocks, &mut |paragraph| {
        inlines(&mut paragraph.inlines, false, apply);
    });
}

/// Inserts a probe at a scoped inline placement.
fn insert_probe(document: &mut Document, target: usize, position: usize, probe: &Inline) {
    let mut count = 0;
    visit_inline_lists(&mut document.blocks, &mut |_| count += 1);
    let target = target % count;
    let mut index = 0;
    visit_inline_lists(&mut document.blocks, &mut |list| {
        if index == target {
            list.insert(position % (list.len() + 1), probe.clone());
        }
        index += 1;
    });
}

/// Physically accepts (current view) or rejects (original view) every
/// revision, as accept-all and reject-all do. Removed content keeps
/// its comment and bookmark markers, which collapse to the removal point.
/// Returns `None` for shapes the reference model does not describe:
/// structural table revisions, and paragraph joins across containers.
fn apply_revisions(document: &Document, view: RevisionView) -> Option<Document> {
    let mut blocks = document.blocks.clone();
    if let Some(Block::Paragraph(paragraph)) = blocks.last_mut() {
        paragraph.mark = None;
    }
    Some(Document {
        blocks: apply_blocks(&blocks, view, false)?,
        section_change: false,
    })
}

fn apply_blocks(blocks: &[Block], view: RevisionView, in_textbox: bool) -> Option<Vec<Block>> {
    let mut output: Vec<Block> = Vec::with_capacity(blocks.len());
    let mut join_next = false;
    for block in blocks {
        let applied = match block {
            Block::Paragraph(paragraph) => {
                let inlines = apply_inlines(&paragraph.inlines, view, false);
                let restores_mark =
                    paragraph.mark_property_change && view == RevisionView::Original;
                let removes_break = !restores_mark
                    && paragraph
                        .mark
                        .is_some_and(|mark| mark.removes_paragraph_break_in(view));
                let applied = Paragraph {
                    mark: None,
                    mark_property_change: false,
                    property_change: false,
                    inlines,
                };
                if join_next {
                    let Some(Block::Paragraph(previous)) = output.last_mut() else {
                        panic!("join_next follows a paragraph");
                    };
                    previous.inlines.extend(applied.inlines);
                } else {
                    output.push(Block::Paragraph(applied));
                }
                join_next = removes_break && !in_textbox;
                continue;
            }
            Block::Table(table) => {
                if table.revision.is_some() {
                    return None;
                }
                Block::Table(Table {
                    revision: None,
                    property_changes: TablePropertyChanges::default(),
                    bookmark: table.bookmark,
                    cells: table
                        .cells
                        .iter()
                        .map(|cell| apply_blocks(cell, view, in_textbox))
                        .collect::<Option<_>>()?,
                })
            }
            Block::Sdt(children) => Block::Sdt(apply_blocks(children, view, in_textbox)?),
            Block::CustomXml(children) => {
                Block::CustomXml(apply_blocks(children, view, in_textbox)?)
            }
            Block::Bookmark(id, children) => {
                Block::Bookmark(*id, apply_blocks(children, view, in_textbox)?)
            }
            Block::Comment(id, children) => {
                Block::Comment(*id, apply_blocks(children, view, in_textbox)?)
            }
        };
        if join_next {
            return None;
        }
        output.push(applied);
    }
    if join_next {
        return None;
    }
    Some(output)
}

fn apply_inlines(inlines: &[Inline], view: RevisionView, removed: bool) -> Vec<Inline> {
    let mut output = Vec::with_capacity(inlines.len());
    for inline in inlines {
        let applied = match inline {
            Inline::Run(run) => {
                if removed {
                    continue;
                }
                let bold = if run.property_change && view == RevisionView::Original {
                    run.prior_bold
                } else {
                    run.bold
                };
                Inline::Run(Run {
                    content: run.content,
                    bold,
                    prior_bold: false,
                    property_change: false,
                })
            }
            Inline::ProofErr => {
                if removed {
                    continue;
                }
                Inline::ProofErr
            }
            Inline::Revision(revision) => {
                let removed = removed || revision.kind.hidden_in(view);
                output.extend(apply_inlines(&revision.children, view, removed));
                continue;
            }
            Inline::Comment(id, children) => {
                Inline::Comment(*id, apply_inlines(children, view, removed))
            }
            Inline::Bookmark(id, children) => {
                Inline::Bookmark(*id, apply_inlines(children, view, removed))
            }
            Inline::Hyperlink(_, children)
            | Inline::Sdt(children)
            | Inline::SmartTag(children)
            | Inline::CustomXml(children)
            | Inline::SimpleField(_, children)
            | Inline::ComplexField(_, children)
                if removed =>
            {
                output.extend(apply_inlines(children, view, removed));
                continue;
            }
            Inline::Hyperlink(target, children) => {
                Inline::Hyperlink(*target, apply_inlines(children, view, removed))
            }
            Inline::Sdt(children) => Inline::Sdt(apply_inlines(children, view, removed)),
            Inline::SmartTag(children) => Inline::SmartTag(apply_inlines(children, view, removed)),
            Inline::CustomXml(children) => {
                Inline::CustomXml(apply_inlines(children, view, removed))
            }
            Inline::SimpleField(target, children) => {
                Inline::SimpleField(*target, apply_inlines(children, view, removed))
            }
            Inline::ComplexField(target, children) => {
                Inline::ComplexField(*target, apply_inlines(children, view, removed))
            }
            Inline::Textbox(blocks) => {
                if removed {
                    continue;
                }
                let Some(blocks) = apply_blocks(blocks, view, true) else {
                    continue;
                };
                Inline::Textbox(blocks)
            }
        };
        output.push(applied);
    }
    output
}

fn require_known_families(projection: &DocumentPackageProjection) -> Result<(), String> {
    let document = &projection.document;
    let facts = &document.structural_facts;
    if document.formatting_completeness
        != (FormattingCompleteness {
            bold: FormattingFactStatus::Known,
            highlight: FormattingFactStatus::Known,
            superscript: FormattingFactStatus::Known,
        })
        || document.revision_status != RevisionProjectionStatus::Complete
        || !matches!(facts.indentation, StructuralFactSet::Known(_))
        || !matches!(facts.numbering_hierarchy, StructuralFactSet::Known(_))
        || !matches!(facts.outline_levels, StructuralFactSet::Known(_))
        || !matches!(facts.bookmarks, StructuralFactSet::Known(_))
        || !matches!(facts.internal_references, StructuralFactSet::Known(_))
        || !matches!(&projection.review_facts.revisions, ReviewFactSet::Known(revisions) if revisions.iter().all(|revision| matches!(revision.content, ReviewDetail::Known(_))))
        || !matches!(&projection.review_facts.comments, ReviewFactSet::Known(comments) if comments.iter().all(|comment| matches!(comment.content, ReviewDetail::Known(_))))
    {
        return Err("generated fact families must be known".to_owned());
    }
    Ok(())
}

/// Compares complete projected facts after physically applying revisions.
fn assert_known_facts_agree(
    left: &DocumentPackageProjection,
    right: &DocumentPackageProjection,
) -> Result<(), String> {
    require_known_families(left)?;
    require_known_families(right)?;
    if left.document.paragraphs != right.document.paragraphs
        || left.document.structural_facts != right.document.structural_facts
        || left.review_facts.comments != right.review_facts.comments
    {
        return Err("applied revisions changed projected facts".to_owned());
    }
    Ok(())
}

fn valid_in_revision(inline: &Inline) -> bool {
    match inline {
        Inline::Hyperlink(..) | Inline::SimpleField(..) | Inline::Textbox(_) => false,
        Inline::Revision(revision) => revision.children.iter().all(valid_in_revision),
        Inline::Comment(_, children)
        | Inline::ComplexField(_, children)
        | Inline::Bookmark(_, children)
        | Inline::Sdt(children)
        | Inline::SmartTag(children)
        | Inline::CustomXml(children) => children.iter().all(valid_in_revision),
        Inline::Run(_) | Inline::ProofErr => true,
    }
}

fn require_scoped_placements(document: &Document) -> Result<(), String> {
    let mut scoped = document.clone();
    scope_document(&mut scoped);
    let actual = write_document(document, &BTreeSet::new());
    let expected = write_document(&scoped, &BTreeSet::new());
    if actual != expected {
        return Err("document contains placements outside the suite".to_owned());
    }
    let mut emitted = BTreeSet::new();
    observe(&actual.0, &mut emitted);
    if emitted.iter().any(|row| EXCLUDED_ROWS.contains(row)) {
        return Err("document emits an excluded context row".to_owned());
    }
    Ok(())
}

proptest! {
    #![proptest_config(config(64))]

    /// Splitting a run into runs with identical formatting changes no fact,
    /// in either direction.
    #[test]
    fn splitting_runs_changes_no_fact(document in document()) {
        let ids = comment_ids(&document);
        let mut split = document.clone();
        map_paragraphs(&mut split.blocks, &mut |paragraph| split_runs(&mut paragraph.inlines));
        for view in VIEWS {
            prop_assert_eq!(project(&document, &ids, view), project(&split, &ids, view));
        }
    }

    /// A proofing marker changes no fact. A zero-width bookmark adds exactly
    /// one bookmark fact and moves no other span.
    #[test]
    fn no_op_markers_change_no_other_fact(
        document in document(),
        target in any::<usize>(),
        position in any::<usize>(),
    ) {
        let ids = comment_ids(&document);
        for view in VIEWS {
            prop_assert!(require_scoped_placements(&document).is_ok());
            let base = project(&document, &ids, view);
            let mut proofed = document.clone();
            insert_probe(&mut proofed, target, position, &Inline::ProofErr);
            prop_assert!(require_scoped_placements(&proofed).is_ok());
            prop_assert_eq!(&base, &project(&proofed, &ids, view));

            let mut bookmarked = document.clone();
            insert_probe(
                &mut bookmarked,
                target,
                position,
                &Inline::Bookmark(PROBE_ID, Vec::new()),
            );
            prop_assert!(require_scoped_placements(&bookmarked).is_ok());
            let mut probed = project(&bookmarked, &ids, view);
            let probe_name = format!("b{PROBE_ID}");
            match (
                &base.document.structural_facts.bookmarks,
                &mut probed.document.structural_facts.bookmarks,
            ) {
                (StructuralFactSet::Known(before), StructuralFactSet::Known(after)) => {
                    let probes = after
                        .iter()
                        .filter(|bookmark| bookmark.name == probe_name)
                        .collect::<Vec<_>>();
                    prop_assert_eq!(probes.len(), 1, "one probe bookmark fact");
                    prop_assert_eq!(probes[0].span.start_utf8, probes[0].span.end_utf8);
                    after.retain(|bookmark| bookmark.name != probe_name);
                    prop_assert_eq!(before, after);
                }
                _ => panic!("probe bookmark facts must be known"),
            }
            probed.document.structural_facts.bookmarks =
                base.document.structural_facts.bookmarks.clone();
            prop_assert_eq!(&base, &probed);
        }
    }

    #[test]
    fn unresolvable_items_preserve_unrelated_known_facts(
        document in document(),
        target in any::<usize>(),
        position in any::<usize>(),
    ) {
        let ids = comment_ids(&document);
        for view in VIEWS {
            let base = project(&document, &ids, view);
            let mut probed_document = document.clone();
            insert_probe(
                &mut probed_document,
                target,
                position,
                &Inline::Revision(Revision {
                    kind: match view {
                        RevisionView::Current => Tracked::Deletion,
                        RevisionView::Original => Tracked::Insertion,
                    },
                    author: Author::Probe,
                    children: vec![Inline::Bookmark(PROBE_ID, Vec::new())],
                }),
            );
            let visible_probe_id = PROBE_ID + 1;
            let mut inserted = false;
            map_paragraphs(&mut probed_document.blocks, &mut |paragraph| {
                if !inserted {
                    paragraph.inlines.push(Inline::Bookmark(visible_probe_id, Vec::new()));
                    inserted = true;
                }
            });
            let (xml, comments) = write_document(&probed_document, &ids);
            let marker = format!(r#"<w:bookmarkEnd w:id="{PROBE_ID}"/>"#);
            let missing_name = format!("missing-{PROBE_ID}");
            let reference_xml = format!(
                r#"{marker}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF {missing_name} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            );
            let xml = xml.replace(&marker, &reference_xml);
            let visible_marker = format!(r#"<w:bookmarkEnd w:id="{visible_probe_id}"/>"#);
            let visible_reference = format!(
                r#"{visible_marker}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF {missing_name} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            );
            let xml = xml.replace(&visible_marker, &visible_reference);
            let mut probed = project_xml(&xml, &comments, view);
            let StructuralFactSet::Known(bookmarks) = &mut probed.document.structural_facts.bookmarks else {
                prop_assert!(false, "hidden bookmark changed the bookmark family's Known status");
                continue;
            };
            let bookmark_probes = bookmarks.iter().filter(|bookmark| bookmark.bookmark_id == u32::try_from(PROBE_ID).unwrap()).collect::<Vec<_>>();
            prop_assert_eq!(bookmark_probes.len(), 1, "the hidden bookmark has one collapsed span");
            prop_assert_eq!(bookmark_probes[0].span.start_utf8, bookmark_probes[0].span.end_utf8);
            bookmarks.retain(|bookmark| ![u32::try_from(PROBE_ID).unwrap(), u32::try_from(visible_probe_id).unwrap()].contains(&bookmark.bookmark_id));
            let StructuralFactSet::Known(references) = &mut probed.document.structural_facts.internal_references else {
                prop_assert!(false, "missing target changed the reference family's Known status");
                continue;
            };
            let reference_count = references.iter().filter(|reference| reference.reference_id == missing_name).count();
            prop_assert_eq!(reference_count, 1, "only the visible missing-target reference preserves its source");
            references.retain(|reference| reference.reference_id != missing_name);
            let ReviewFactSet::Known(revisions) = &mut probed.review_facts.revisions else {
                prop_assert!(false, "hidden annotations changed the revision family's Known status");
                continue;
            };
            revisions.retain(|revision| revision.author != PROBE_AUTHOR);
            prop_assert_eq!(&base, &probed);
        }
    }

    /// The current view knows what accepting every revision produces, and
    /// the original view what rejecting every revision produces.
    #[test]
    fn views_agree_with_physically_applied_revisions(document in document()) {
        let ids = comment_ids(&document);
        for view in VIEWS {
            let applied = apply_revisions(&document, view).expect("generated placements have a complete applied model");
            let projected = project(&document, &ids, view);
            let reference = project(&applied, &ids, RevisionView::Current);
            prop_assert!(known_revisions(&reference).is_empty(), "applying leaves no revision");
            let agreed = assert_known_facts_agree(&projected, &reference);
            prop_assert!(agreed.is_ok(), "{:?}: {}", view, agreed.unwrap_err());
        }
    }

    /// Wrapping content in a tracked insertion adds exactly one revision,
    /// located over that content, and changes nothing else in the current
    /// view.
    #[test]
    fn wrapping_content_in_an_insertion_adds_one_covering_revision(
        document in document(),
        paragraph_index in any::<usize>(),
        start in any::<usize>(),
        length in any::<usize>(),
    ) {
        let paragraphs = document
            .blocks
            .iter()
            .enumerate()
            .filter_map(|(index, block)| matches!(block, Block::Paragraph(_)).then_some(index))
            .collect::<Vec<_>>();
        let block_index = *paragraphs.get(paragraph_index % paragraphs.len()).expect("the scoped document includes a paragraph");
        let mut wrapped = document.clone();
        let Block::Paragraph(paragraph) = &mut wrapped.blocks[block_index] else {
            panic!("selected a paragraph block");
        };
        let start = start % (paragraph.inlines.len() + 1);
        let mut end = start;
        while end < paragraph.inlines.len()
            && end - start < length % 4
            && valid_in_revision(&paragraph.inlines[end])
        {
            end += 1;
        }
        let content = paragraph.inlines.drain(start..end).collect::<Vec<_>>();
        paragraph.inlines.insert(
            start,
            Inline::Revision(Revision {
                kind: Tracked::Insertion,
                author: Author::Probe,
                children: content,
            }),
        );
        prop_assert!(require_scoped_placements(&wrapped).is_ok());
        let ids = comment_ids(&document);
        let view = RevisionView::Current;
        let base = project(&document, &ids, view);
        let mut probed = project(&wrapped, &ids, view);
        let ReviewFactSet::Known(probed_revisions) = &mut probed.review_facts.revisions else {
            panic!("revision facts are unknown");
        };
        let probes = probed_revisions
            .iter()
            .enumerate()
            .filter(|(_, revision)| revision.author == PROBE_AUTHOR)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        prop_assert_eq!(probes.len(), 1, "exactly one probe revision");
        let probe = probed_revisions.remove(probes[0]);
        let expected = Model::build(&wrapped, view);
        let (start_point, end_point, payload) = expected.revisions[probes[0]]
            .located
            .clone()
            .expect("an inline revision has a modelled location");
        prop_assert_eq!(
            probe.content,
            ReviewDetail::Known(RevisionContent {
                span: expected.span((start_point, end_point)),
                payload,
            })
        );
        prop_assert_eq!(&base, &probed);
    }
}

// ---------------------------------------------------------------------------

fn single(document_body: &str, view: RevisionView) -> DocumentPackageProjection {
    let xml = format!(r#"<w:document xmlns:w="{W}"><w:body>{document_body}</w:body></w:document>"#);
    project_xml(&xml, &format!(r#"<w:comments xmlns:w="{W}"/>"#), view)
}

#[test]
fn moved_paragraph_marks_join_like_deleted_and_inserted_marks() {
    for (mark, view, expected) in [
        ("del", RevisionView::Current, &["ab"][..]),
        ("del", RevisionView::Original, &["a", "b"]),
        ("moveFrom", RevisionView::Current, &["ab"]),
        ("moveFrom", RevisionView::Original, &["a", "b"]),
        ("ins", RevisionView::Current, &["a", "b"]),
        ("ins", RevisionView::Original, &["ab"]),
        ("moveTo", RevisionView::Current, &["a", "b"]),
        ("moveTo", RevisionView::Original, &["ab"]),
    ] {
        let projection = single(
            &format!(
                r#"<w:p><w:pPr><w:rPr><w:{mark} w:id="1" w:author="A"/></w:rPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>"#
            ),
            view,
        );
        let texts = projection
            .document
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.text.as_str())
            .collect::<Vec<_>>();
        assert_eq!(texts, expected, "{mark} in {view:?}");
    }
}

#[test]
fn endnote_references_materialize_like_footnote_references() {
    for view in VIEWS {
        let projection = single(
            r#"<w:p><w:r><w:t>a</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r><w:r><w:t>z</w:t></w:r></w:p>"#,
            view,
        );
        assert_eq!(projection.document.paragraphs[0].text, "a\u{0002}\u{0002}z");
    }
}

#[test]
fn property_change_revisions_are_located() {
    let projection = single(
        &format!(
            r#"<w:p><w:pPr><w:pPrChange {TRACKED}><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>ab</w:t></w:r></w:p>"#
        ),
        RevisionView::Current,
    );
    for revision in known_revisions(&projection) {
        assert!(
            matches!(revision.content, ReviewDetail::Known(_)),
            "{:?} has no location",
            revision.kind
        );
    }
}

#[test]
fn paragraph_joins_keep_bookmark_facts() {
    let suffix = r#"<w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>b</w:t></w:r><w:bookmarkEnd w:id="0"/><w:hyperlink w:anchor="b"/><w:fldSimple w:instr="REF b"/>"#;
    let projection = single(
        &format!(
            r"<w:p><w:pPr><w:rPr><w:del {TRACKED}/></w:rPr></w:pPr><w:r><w:t>é😀</w:t></w:r></w:p><w:p>{suffix}</w:p>"
        ),
        RevisionView::Current,
    );
    let equivalent = single(
        &format!(r"<w:p><w:r><w:t>é😀</w:t></w:r>{suffix}</w:p>"),
        RevisionView::Current,
    );
    assert_eq!(
        projection.document.structural_facts.bookmarks,
        equivalent.document.structural_facts.bookmarks
    );
    assert_eq!(
        projection.document.structural_facts.internal_references,
        equivalent.document.structural_facts.internal_references
    );
    let StructuralFactSet::Known(bookmarks) = &projection.document.structural_facts.bookmarks
    else {
        panic!("joined bookmark facts must be Known");
    };
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].paragraph_ordinal, 0);
    assert_eq!(
        (bookmarks[0].span.start_utf8, bookmarks[0].span.end_utf8),
        (6, 7)
    );
    assert_eq!(
        (bookmarks[0].span.start_utf16, bookmarks[0].span.end_utf16),
        (3, 4)
    );
    let StructuralFactSet::Known(references) =
        &projection.document.structural_facts.internal_references
    else {
        panic!("joined reference facts must be Known");
    };
    assert_eq!(references.len(), 3, "two sources and one target");
    assert!(
        references
            .iter()
            .all(|reference| reference.paragraph_ordinal == 0)
    );
}

#[test]
fn bookmarks_in_hidden_content_keep_bookmark_facts() {
    let projection = single(
        &format!(
            r#"<w:p><w:r><w:t>a</w:t></w:r><w:del {TRACKED}><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:delText>b</w:delText></w:r><w:bookmarkEnd w:id="0"/></w:del></w:p>"#
        ),
        RevisionView::Current,
    );
    assert!(
        matches!(
            projection.document.structural_facts.bookmarks,
            StructuralFactSet::Known(_)
        ),
        "{:?}",
        projection.document.structural_facts.bookmarks
    );
}

#[test]
fn bookmarks_in_block_content_controls_keep_bookmark_facts() {
    let content = r#"<w:bookmarkStart w:id="0" w:name="b"/><w:p><w:r><w:t>a</w:t></w:r></w:p><w:bookmarkEnd w:id="0"/>"#;
    for body in [
        format!("<w:sdt><w:sdtContent>{content}</w:sdtContent></w:sdt>"),
        format!("<w:customXml>{content}</w:customXml>"),
    ] {
        let projection = single(&body, RevisionView::Current);
        let reference = single(content, RevisionView::Current);
        assert_eq!(
            projection.document.structural_facts.bookmarks,
            reference.document.structural_facts.bookmarks
        );
        assert!(matches!(
            projection.document.structural_facts.bookmarks,
            StructuralFactSet::Known(_)
        ));
    }
}

#[test]
fn references_in_hidden_content_keep_internal_references() {
    let projection = single(
        &format!(
            r#"<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id="0"/><w:del {TRACKED}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText> REF b \h </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del></w:p>"#
        ),
        RevisionView::Current,
    );
    assert!(
        matches!(
            projection.document.structural_facts.internal_references,
            StructuralFactSet::Known(_)
        ),
        "{:?}",
        projection.document.structural_facts.internal_references
    );
}

#[test]
fn paragraph_joins_coalesce_formatting_spans() {
    let projection = single(
        &format!(
            r"<w:p><w:pPr><w:rPr><w:del {TRACKED}/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>"
        ),
        RevisionView::Current,
    );
    assert_eq!(
        projection.document.paragraphs[0].formatting,
        [TextFormattingSpan {
            start_utf16: 0,
            end_utf16: 2,
            style: TextStyle::Bold,
        }],
        "a joined paragraph reports what an unjoined paragraph with the same runs reports"
    );
}

const TRACKED: &str = r#"w:id="1" w:author="A""#;

#[derive(Clone, Debug)]
struct TablePropertyDocument {
    rows: Vec<Vec<Vec<&'static str>>>,
    snapshots: [bool; 5],
    placement: TablePlacement,
    namespace: NamespaceProfile,
}

#[derive(Clone, Copy, Debug)]
enum NamespaceProfile {
    Transitional,
    Strict,
}

#[derive(Clone, Copy, Debug)]
enum TablePlacement {
    Body,
    ContentControl,
    CustomXml,
}

fn table_property_document() -> impl Strategy<Value = TablePropertyDocument> {
    (
        collection::vec(
            collection::vec(
                collection::vec(sample::select(vec!["", "ab", "é", "😀", "a😀é"]), 1..4),
                1..4,
            ),
            1..4,
        ),
        any::<[bool; 5]>(),
        sample::select(vec![
            TablePlacement::Body,
            TablePlacement::ContentControl,
            TablePlacement::CustomXml,
        ]),
        sample::select(vec![
            NamespaceProfile::Transitional,
            NamespaceProfile::Strict,
        ]),
    )
        .prop_map(
            |(rows, snapshots, placement, namespace)| TablePropertyDocument {
                rows,
                snapshots,
                placement,
                namespace,
            },
        )
}

impl TablePropertyDocument {
    fn xml(&self) -> String {
        let namespace = match self.namespace {
            NamespaceProfile::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
            NamespaceProfile::Transitional => W,
        };
        let mut markup = Markup {
            xml: format!(r#"<w:document xmlns:w="{namespace}"><w:body>"#),
            comment_ids: BTreeSet::new(),
            next_revision: 1,
        };
        markup.paragraph(&empty_paragraph());
        let (opening, closing) = match self.placement {
            TablePlacement::Body => ("", ""),
            TablePlacement::ContentControl => ("<w:sdt><w:sdtContent>", "</w:sdtContent></w:sdt>"),
            TablePlacement::CustomXml => ("<w:customXml>", "</w:customXml>"),
        };
        markup.xml.push_str(opening);
        markup
            .xml
            .push_str(r#"<w:tbl><w:tblPr><w:tblStyle w:val="MissingTableStyle"/>"#);
        if self.snapshots[0] {
            markup.change("tblPrChange", "<w:tblPr/>");
        }
        markup.xml.push_str("</w:tblPr><w:tblGrid>");
        if self.snapshots[1] {
            markup.change("tblGridChange", "<w:tblGrid/>");
        }
        markup.xml.push_str("</w:tblGrid>");
        for row in &self.rows {
            markup.xml.push_str("<w:tr>");
            if self.snapshots[4] {
                markup.xml.push_str("<w:tblPrEx>");
                markup.change("tblPrExChange", "<w:tblPrEx/>");
                markup.xml.push_str("</w:tblPrEx>");
            }
            markup.xml.push_str("<w:trPr>");
            if self.snapshots[2] {
                markup.change("trPrChange", "<w:trPr/>");
            }
            markup.xml.push_str("</w:trPr>");
            for cell in row {
                markup.xml.push_str("<w:tc><w:tcPr>");
                if self.snapshots[3] {
                    markup.change("tcPrChange", "<w:tcPr/>");
                }
                markup.xml.push_str("</w:tcPr>");
                for text in cell {
                    markup.paragraph(&Paragraph {
                        inlines: vec![Inline::Run(Run {
                            content: RunContent::Text(text),
                            bold: false,
                            prior_bold: false,
                            property_change: false,
                        })],
                        ..empty_paragraph()
                    });
                }
                markup.xml.push_str("</w:tc>");
            }
            markup.xml.push_str("</w:tr>");
        }
        markup.xml.push_str("</w:tbl>");
        markup.xml.push_str(closing);
        markup.paragraph(&empty_paragraph());
        markup.xml.push_str("</w:body></w:document>");
        markup.xml
    }

    #[allow(clippy::too_many_lines)] // Compare every generated owner in one traversal.
    fn check(&self, view: RevisionView) -> Result<(), String> {
        let projection = project_xml(
            &self.xml(),
            &format!(r#"<w:comments xmlns:w="{W}"/>"#),
            view,
        );
        let mut texts = vec![""];
        let mut structures = vec![None];
        let mut owners = Vec::new();
        let point = |ordinal: usize, text: &str| ReviewPoint {
            paragraph_ordinal: ordinal,
            utf8: u32::try_from(text.len()).unwrap(),
            utf16: u32::try_from(text.encode_utf16().count()).unwrap(),
        };
        for (row_index, row) in self.rows.iter().enumerate() {
            let row_start = point(texts.len(), "");
            let mut cells = Vec::new();
            for (column, cell) in row.iter().enumerate() {
                let start = point(texts.len(), "");
                for text in cell {
                    texts.push(text);
                    structures.push(Some(ParagraphStructure {
                        table_ordinal: 0,
                        row: row_index,
                        column,
                    }));
                }
                cells.push(ReviewSpan {
                    start,
                    end: point(texts.len() - 1, texts[texts.len() - 1]),
                });
            }
            owners.push((
                ReviewSpan {
                    start: row_start,
                    end: point(texts.len() - 1, texts[texts.len() - 1]),
                },
                cells,
            ));
        }
        let table_span = ReviewSpan {
            start: point(1, ""),
            end: point(texts.len() - 1, texts[texts.len() - 1]),
        };
        texts.push("");
        structures.push(None);
        let mut expected = Vec::new();
        if self.snapshots[0] {
            expected.push((RevisionFactKind::TablePropertiesChange, table_span));
        }
        if self.snapshots[1] {
            expected.push((RevisionFactKind::TableGridChange, table_span));
        }
        for (row, cells) in owners {
            if self.snapshots[4] {
                expected.push((RevisionFactKind::TablePropertiesExceptionChange, row));
            }
            if self.snapshots[2] {
                expected.push((RevisionFactKind::TableRowPropertiesChange, row));
            }
            if self.snapshots[3] {
                for cell in cells {
                    expected.push((RevisionFactKind::TableCellPropertiesChange, cell));
                }
            }
        }
        let paragraphs = &projection.document.paragraphs;
        if paragraphs.iter().enumerate().any(|(ordinal, paragraph)| {
            paragraph.ordinal != ordinal
                || paragraph.style_id.is_some()
                || paragraph.package_paragraph_id.is_some()
                || paragraph.alignment.is_some()
                || !paragraph.formatting.is_empty()
        }) {
            return Err("table paragraph ordinals or direct formatting differ".to_owned());
        }
        if paragraphs
            .iter()
            .map(|p| p.text.as_str())
            .collect::<Vec<_>>()
            != texts
            || paragraphs
                .iter()
                .map(|p| p.structure.clone())
                .collect::<Vec<_>>()
                != structures
        {
            return Err("table text or coordinates differ from the generated cells".to_owned());
        }
        let empty_structure = DocumentStructureFacts {
            bookmarks: StructuralFactSet::Known(Vec::new()),
            internal_references: StructuralFactSet::Known(Vec::new()),
            indentation: StructuralFactSet::Known(Vec::new()),
            numbering_hierarchy: StructuralFactSet::Known(Vec::new()),
            outline_levels: StructuralFactSet::Known(Vec::new()),
        };
        if projection.document.structural_facts != empty_structure
            || projection.review_facts.comments != ReviewFactSet::Known(Vec::new())
        {
            return Err("table annotations and paragraph properties are empty".to_owned());
        }
        if projection.document.revision_status != RevisionProjectionStatus::Complete
            || projection.document.formatting_completeness
                != (FormattingCompleteness {
                    bold: FormattingFactStatus::Known,
                    highlight: FormattingFactStatus::Known,
                    superscript: FormattingFactStatus::Known,
                })
        {
            return Err(
                "table projection statuses differ from their declared semantics".to_owned(),
            );
        }
        let ReviewFactSet::Known(revisions) = &projection.review_facts.revisions else {
            return Err("bounded table revisions are unknown".to_owned());
        };
        if revisions.len() != expected.len() {
            return Err("table revision cardinality differs".to_owned());
        }
        for (index, (revision, (kind, span))) in revisions.iter().zip(expected).enumerate() {
            if revision.kind != kind
                || revision.author != format!("Author {}", (index + 1) % 3)
                || revision.date.as_deref()
                    != Some(format!("2024-01-0{}T00:00:00Z", (index + 1) % 9 + 1).as_str())
                || revision.revision_id.as_deref() != Some((index + 1).to_string().as_str())
                || revision.content
                    != ReviewDetail::Known(RevisionContent {
                        span,
                        payload: RevisionPayload::FormattingOnly,
                    })
            {
                return Err(format!(
                    "table revision {index} differs from its owner span"
                ));
            }
        }
        Ok(())
    }
}

proptest! {
    #![proptest_config(config(64))]
    #[test]
    fn table_property_revisions_match_their_owners(document in table_property_document()) {
        for view in VIEWS {
            let checked = document.check(view);
            prop_assert!(checked.is_ok(), "{}", checked.unwrap_err());
        }
    }
}

#[test]
fn table_property_revisions_cover_distinct_rows_and_cells() {
    let document = TablePropertyDocument {
        rows: vec![
            vec![vec!["", "é😀"], vec!["a"]],
            vec![vec!["b", "😀"], vec![""]],
        ],
        snapshots: [true; 5],
        placement: TablePlacement::Body,
        namespace: NamespaceProfile::Transitional,
    };
    for view in VIEWS {
        document.check(view).unwrap();
    }
}
