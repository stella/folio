use std::collections::{BTreeSet, HashMap, HashSet};

use quick_xml::XmlVersion;
use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::ProjectionError;
use crate::projection::compatibility::{CompatibilityAction, MarkupCompatibility};
use crate::projection::namespaces::OoxmlNamespace;
use crate::projection::review::{
    AttributedRevision, ReviewDetail, ReviewFactSet, ReviewFactUnknownReason, ReviewPoint,
    ReviewSpan, RevisionContent, RevisionFactKind,
};
use crate::projection::structure::{
    ParagraphProperties, RawBlockPoint, RawBookmarkRange, RawInternalReference,
    StructuralFactUnknownReason,
};
use crate::projection::styles::{
    parse_indentation, parse_level_attribute, parse_outline_level_attribute, parse_u32_attribute,
};

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct PackageParagraphId(u32);

impl PackageParagraphId {
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return None;
        }
        let value = u32::from_str_radix(value, 16).ok()?;
        (value != 0 && i32::try_from(value).is_ok()).then_some(Self(value))
    }

    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub enum TextStyle {
    Bold,
    Highlight,
    Superscript,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextFormattingSpan {
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub style: TextStyle,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphStructure {
    pub table_ordinal: usize,
    pub row: usize,
    pub column: usize,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum RevisionView {
    #[default]
    Current,
    Original,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum TextMaterialization {
    #[default]
    WordHost,
    ReadablePlainText,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub enum RevisionUnsupportedReason {
    IncompatibleParagraphMerge,
    StructuralTableRevision,
    UnsupportedRevisionMarkup,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum RevisionProjectionStatus {
    Complete,
    Incomplete(Vec<RevisionUnsupportedReason>),
}

impl RevisionProjectionStatus {
    #[must_use]
    pub const fn is_complete(&self) -> bool {
        matches!(self, Self::Complete)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ParagraphMarkRevision {
    Insertion,
    Deletion,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum TextControl {
    Tab,
    LineBreak,
    PageBreak,
    ColumnBreak,
    CarriageReturn,
    FootnoteReference,
    SoftHyphen,
    NoBreakHyphen,
}

impl TextControl {
    pub(super) const fn materialize(
        self,
        materialization: TextMaterialization,
    ) -> Option<&'static str> {
        match materialization {
            TextMaterialization::WordHost => Some(match self {
                Self::Tab => "\t",
                Self::LineBreak => "\u{000b}",
                Self::PageBreak => "\u{000c}",
                Self::ColumnBreak => "\u{000e}",
                Self::CarriageReturn => "\r",
                Self::FootnoteReference => "\u{0002}",
                Self::SoftHyphen => "\u{001f}",
                Self::NoBreakHyphen => "\u{001e}",
            }),
            TextMaterialization::ReadablePlainText => match self {
                Self::Tab => Some("\t"),
                Self::LineBreak | Self::PageBreak | Self::ColumnBreak | Self::CarriageReturn => {
                    Some("\n")
                }
                Self::FootnoteReference => None,
                Self::SoftHyphen => Some("\u{00ad}"),
                Self::NoBreakHyphen => Some("\u{2011}"),
            },
        }
    }
}

pub(super) struct RawProjectedParagraph {
    pub ordinal: usize,
    pub package_paragraph_id: Option<PackageParagraphId>,
    pub text: String,
    utf16_len: u32,
    pub formatting: Vec<TextFormattingSpan>,
    pub structure: Option<ParagraphStructure>,
    pub properties: ParagraphProperties,
}

pub(super) struct RawDocumentProjection {
    pub paragraphs: Vec<RawProjectedParagraph>,
    pub bookmarks: Result<Vec<RawBookmarkRange>, StructuralFactUnknownReason>,
    pub references: Result<Vec<RawInternalReference>, StructuralFactUnknownReason>,
    pub revision_status: RevisionProjectionStatus,
    pub review_revisions: Option<ReviewFactSet<AttributedRevision>>,
    pub review_comment_anchors: HashMap<String, ReviewSpan>,
}

struct ParagraphBuilder {
    package_paragraph_id: Option<PackageParagraphId>,
    text: String,
    utf16_len: u32,
    formatting: Vec<TextFormattingSpan>,
    structure: Option<ParagraphStructure>,
    properties: ParagraphProperties,
    paragraph_mark_revision: Option<ParagraphMarkRevision>,
}

impl ParagraphBuilder {
    fn append(&mut self, text: &str, styles: DirectTextStyles) -> Result<(), ProjectionError> {
        let units = u32::try_from(text.encode_utf16().count())
            .map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let end = self
            .utf16_len
            .checked_add(units)
            .ok_or(ProjectionError::InvalidDocumentXml)?;
        if styles.bold {
            self.append_style(end, TextStyle::Bold);
        }
        if styles.highlighted {
            self.append_style(end, TextStyle::Highlight);
        }
        if styles.superscript {
            self.append_style(end, TextStyle::Superscript);
        }
        self.text.push_str(text);
        self.utf16_len = end;
        Ok(())
    }

    fn append_style(&mut self, end_utf16: u32, style: TextStyle) {
        if self.utf16_len >= end_utf16 {
            return;
        }
        if let Some(previous) = self
            .formatting
            .iter_mut()
            .rev()
            .find(|span| span.style == style)
            && previous.end_utf16 == self.utf16_len
        {
            previous.end_utf16 = end_utf16;
            return;
        }
        self.formatting.push(TextFormattingSpan {
            start_utf16: self.utf16_len,
            end_utf16,
            style,
        });
    }

    fn truncate(&mut self, utf8_len: usize, utf16_len: u32) {
        self.text.truncate(utf8_len);
        self.utf16_len = utf16_len;
        self.formatting.retain(|span| span.start_utf16 < utf16_len);
        for span in &mut self.formatting {
            span.end_utf16 = span.end_utf16.min(utf16_len);
        }
    }
}

#[derive(Clone, Copy, Default)]
struct DirectTextStyles {
    bold: bool,
    highlighted: bool,
    superscript: bool,
}

struct RunFrame {
    text: String,
    styles: DirectTextStyles,
    hidden: bool,
    direct_child_count: usize,
}

struct RunPropertiesFrame {
    run_frame: usize,
    hidden_eligible: bool,
}

struct TableFrame {
    ordinal: usize,
    next_row: usize,
}

struct RowFrame {
    table_ordinal: usize,
    row: usize,
    next_column: usize,
}

#[derive(Clone, Copy)]
struct CellFrame {
    table_ordinal: usize,
    row: usize,
    column: usize,
}

struct SdtFrame {
    placeholder: bool,
    text_utf8_start: usize,
    text_utf16_start: u32,
}

#[derive(Clone, Copy)]
enum PseudoTextKind {
    Text { preserve_space: bool },
    Instruction,
}

struct PseudoTextFrame {
    kind: PseudoTextKind,
    text: String,
}

impl PseudoTextFrame {
    const fn text(preserve_space: bool) -> Self {
        Self {
            kind: PseudoTextKind::Text { preserve_space },
            text: String::new(),
        }
    }

    const fn instruction() -> Self {
        Self {
            kind: PseudoTextKind::Instruction,
            text: String::new(),
        }
    }
}

struct FieldFrame {
    source: Option<RawBlockPoint>,
    instruction: String,
    separated: bool,
}

struct RevisionFrame {
    review_index: Option<usize>,
    start: Option<ReviewPoint>,
    suppressed: bool,
    text: String,
}

#[derive(Default)]
enum ReviewRevisionCollection {
    #[default]
    Disabled,
    Complete {
        maximum_facts: usize,
        remaining_detail_bytes: usize,
        revisions: Vec<AttributedRevision>,
    },
    LimitExceeded,
}

#[derive(Clone, Copy)]
pub(super) struct ReviewProjectionLimits {
    pub maximum_facts: usize,
    pub maximum_detail_bytes: usize,
}

enum Frame {
    Other,
    Body,
    Table(TableFrame),
    Row(RowFrame),
    Cell(CellFrame),
    Paragraph,
    Run(RunFrame),
    RunProperties(RunPropertiesFrame),
    ParagraphProperties,
    TableRowProperties,
    TableCellProperties,
    NumberingProperties,
    Hyperlink(Option<RawInternalReference>),
    Revision(RevisionFrame),
    ChangeSnapshot,
    Textbox,
    Sdt(SdtFrame),
    PseudoText(PseudoTextFrame),
}

#[allow(clippy::too_many_lines)] // A single streaming event loop keeps XML parsing one-pass.
pub(super) fn project_document_xml(
    xml: &[u8],
    maximum_paragraphs: usize,
    revision_view: RevisionView,
    text_materialization: TextMaterialization,
    review_limits: Option<ReviewProjectionLimits>,
) -> Result<RawDocumentProjection, ProjectionError> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut state = ProjectionState {
        maximum_paragraphs,
        revision_view,
        text_materialization,
        bookmarks_complete: true,
        references_complete: true,
        review_revisions: review_limits.map_or(ReviewRevisionCollection::Disabled, |limits| {
            ReviewRevisionCollection::Complete {
                maximum_facts: limits.maximum_facts,
                remaining_detail_bytes: limits.maximum_detail_bytes,
                revisions: Vec::new(),
            }
        }),
        ..ProjectionState::default()
    };
    let mut compatibility = MarkupCompatibility::default();
    loop {
        match reader
            .read_event()
            .map_err(|_| ProjectionError::InvalidDocumentXml)?
        {
            Event::Start(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                let namespace = OoxmlNamespace::from_resolved(&namespace);
                if compatibility.start(&reader, namespace, &element)?
                    == CompatibilityAction::Process
                {
                    state.start(&reader, namespace, &element)?;
                }
            }
            Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                let namespace = OoxmlNamespace::from_resolved(&namespace);
                if compatibility.empty(&reader, namespace, &element)?
                    == CompatibilityAction::Process
                {
                    state.start(&reader, namespace, &element)?;
                    state.end()?;
                }
            }
            Event::Text(text) => {
                if !compatibility.is_suppressed()
                    && matches!(state.frames.last(), Some(Frame::PseudoText(_)))
                {
                    let decoded = text
                        .xml10_content()
                        .map_err(|_| ProjectionError::InvalidDocumentXml)?;
                    state.append_pseudo_content(&decoded)?;
                }
            }
            Event::CData(text) => {
                if !compatibility.is_suppressed()
                    && matches!(state.frames.last(), Some(Frame::PseudoText(_)))
                {
                    let decoded = text
                        .xml10_content()
                        .map_err(|_| ProjectionError::InvalidDocumentXml)?;
                    state.append_pseudo_content(&decoded)?;
                }
            }
            Event::GeneralRef(reference) => {
                if !compatibility.is_suppressed()
                    && matches!(state.frames.last(), Some(Frame::PseudoText(_)))
                {
                    let resolved_character = reference
                        .resolve_char_ref()
                        .map_err(|_| ProjectionError::InvalidDocumentXml)?;
                    let resolved = if let Some(character) = resolved_character {
                        character.to_string()
                    } else {
                        let name = reference
                            .decode()
                            .map_err(|_| ProjectionError::InvalidDocumentXml)?;
                        resolve_predefined_entity(&name)
                            .ok_or(ProjectionError::InvalidDocumentXml)?
                            .to_owned()
                    };
                    state.append_pseudo_content(&resolved)?;
                }
            }
            Event::End(element) => {
                let (namespace, local_name) = reader.resolver().resolve_element(element.name());
                if compatibility.end(
                    OoxmlNamespace::from_resolved(&namespace),
                    local_name.as_ref(),
                )? == CompatibilityAction::Process
                {
                    state.end()?;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !state.body_seen {
        return Err(ProjectionError::MissingDocumentBody);
    }
    if state.current_paragraph.is_some() || !state.frames.is_empty() || !compatibility.is_complete()
    {
        return Err(ProjectionError::InvalidDocumentXml);
    }
    let paragraph_merges = if state.paragraph_mark_revisions.is_empty() {
        false
    } else {
        normalize_paragraph_revision_view(
            &mut state.paragraphs,
            &mut state.paragraph_mark_revisions,
            state.revision_view,
            &mut state.revision_unsupported,
        )?
    };
    if paragraph_merges {
        state.bookmarks_complete = false;
        state.references_complete = false;
        state.review_comment_anchors.clear();
        if let ReviewRevisionCollection::Complete { revisions, .. } = &mut state.review_revisions {
            for revision in revisions {
                revision.content =
                    ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation);
            }
        }
    }
    let bookmarks = if state.bookmarks_complete
        && state.open_bookmarks.is_empty()
        && (state.bookmarks.is_empty() || !state.paragraphs.is_empty())
    {
        Ok(state.bookmarks)
    } else {
        Err(StructuralFactUnknownReason::IncompleteBookmarkRanges)
    };
    let references = if state.references_complete && state.fields.is_empty() {
        Ok(state.references)
    } else {
        Err(StructuralFactUnknownReason::UnsupportedInternalReferences)
    };
    Ok(RawDocumentProjection {
        paragraphs: state.paragraphs,
        bookmarks,
        references,
        revision_status: if state.revision_unsupported.is_empty() {
            RevisionProjectionStatus::Complete
        } else {
            RevisionProjectionStatus::Incomplete(state.revision_unsupported.into_iter().collect())
        },
        review_revisions: match state.review_revisions {
            ReviewRevisionCollection::Disabled => None,
            ReviewRevisionCollection::Complete { revisions, .. } => {
                Some(ReviewFactSet::Known(revisions))
            }
            ReviewRevisionCollection::LimitExceeded => Some(ReviewFactSet::Unknown(
                ReviewFactUnknownReason::ResourceLimit,
            )),
        },
        review_comment_anchors: state.review_comment_anchors,
    })
}

#[derive(Default)]
struct ProjectionState {
    frames: Vec<Frame>,
    paragraphs: Vec<RawProjectedParagraph>,
    current_paragraph: Option<ParagraphBuilder>,
    body_seen: bool,
    next_table_ordinal: usize,
    maximum_paragraphs: usize,
    open_bookmarks: HashMap<u32, (String, RawBlockPoint)>,
    seen_bookmark_ids: HashSet<u32>,
    bookmarks: Vec<RawBookmarkRange>,
    references: Vec<RawInternalReference>,
    bookmarks_complete: bool,
    references_complete: bool,
    fields: Vec<FieldFrame>,
    paragraph_mark_revisions: HashMap<usize, ParagraphMarkRevision>,
    revision_view: RevisionView,
    text_materialization: TextMaterialization,
    revision_unsupported: BTreeSet<RevisionUnsupportedReason>,
    review_revisions: ReviewRevisionCollection,
    open_review_comment_anchors: HashMap<String, ReviewPoint>,
    review_comment_anchors: HashMap<String, ReviewSpan>,
    invalid_review_comment_anchors: HashSet<String>,
}

impl ProjectionState {
    #[allow(clippy::too_many_lines)] // The frame transition table is one streaming state machine.
    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: OoxmlNamespace,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let local_name = element.local_name();
        let name = if namespace == OoxmlNamespace::Wordprocessing {
            local_name.as_ref()
        } else {
            b""
        };
        let direct_run_child = if let Some(Frame::Run(run)) = self.frames.last_mut() {
            let child = run.direct_child_count;
            run.direct_child_count = run
                .direct_child_count
                .checked_add(1)
                .ok_or(ProjectionError::InvalidDocumentXml)?;
            let frame_index = self
                .frames
                .len()
                .checked_sub(1)
                .ok_or(ProjectionError::InvalidDocumentXml)?;
            Some((frame_index, child))
        } else {
            None
        };

        if name == b"body" {
            if self.body_seen {
                return Err(ProjectionError::InvalidDocumentXml);
            }
            self.body_seen = true;
            self.frames.push(Frame::Body);
            return Ok(());
        }
        if !self.inside_body() {
            self.frames.push(Frame::Other);
            return Ok(());
        }
        if self.inside_change_snapshot() {
            self.frames.push(Frame::Other);
            return Ok(());
        }
        let attributed_revision = self.record_attributed_revision(reader, element, name)?;
        if name == b"txbxContent" {
            self.frames.push(Frame::Textbox);
            return Ok(());
        }
        if self.inside_textbox() {
            if matches!(
                name,
                b"commentRangeStart" | b"commentRangeEnd" | b"commentReference"
            ) {
                self.suppress_review_comment_anchor(reader, element)?;
            }
            self.frames.push(Frame::Other);
            return Ok(());
        }
        match name {
            b"commentRangeStart" => self.start_review_comment_anchor(reader, element)?,
            b"commentRangeEnd" => self.end_review_comment_anchor(reader, element)?,
            b"commentReference" => self.record_review_comment_reference(reader, element)?,
            _ => {}
        }
        if is_change_snapshot(name) {
            if self.revision_view == RevisionView::Original {
                self.revision_unsupported
                    .insert(RevisionUnsupportedReason::UnsupportedRevisionMarkup);
            }
            self.frames.push(Frame::ChangeSnapshot);
            return Ok(());
        }
        if is_unsupported_revision_markup(name) {
            self.revision_unsupported
                .insert(RevisionUnsupportedReason::UnsupportedRevisionMarkup);
            self.frames.push(Frame::Other);
            return Ok(());
        }

        let frame = match name {
            b"tbl" => {
                let ordinal = self.next_table_ordinal;
                self.next_table_ordinal = self
                    .next_table_ordinal
                    .checked_add(1)
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                Frame::Table(TableFrame {
                    ordinal,
                    next_row: 0,
                })
            }
            b"tr" => {
                let Some(table) = enclosing_table(&mut self.frames) else {
                    self.frames.push(Frame::Other);
                    return Ok(());
                };
                let row = table.next_row;
                table.next_row = table
                    .next_row
                    .checked_add(1)
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                Frame::Row(RowFrame {
                    table_ordinal: table.ordinal,
                    row,
                    next_column: 0,
                })
            }
            b"tc" => {
                let Some(row) = enclosing_row(&mut self.frames) else {
                    self.frames.push(Frame::Other);
                    return Ok(());
                };
                let column = row.next_column;
                row.next_column = row
                    .next_column
                    .checked_add(1)
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                Frame::Cell(CellFrame {
                    table_ordinal: row.table_ordinal,
                    row: row.row,
                    column,
                })
            }
            b"p" => {
                if self.current_paragraph.is_some() {
                    self.frames.push(Frame::Other);
                    return Ok(());
                }
                if self.paragraphs.len() >= self.maximum_paragraphs {
                    return Err(ProjectionError::TooManyParagraphs);
                }
                let package_paragraph_id = attribute_2010(reader, element, b"paraId")?
                    .as_deref()
                    .and_then(PackageParagraphId::parse);
                let structure = self.frames.iter().rev().find_map(|frame| match frame {
                    Frame::Cell(cell) => Some(ParagraphStructure {
                        table_ordinal: cell.table_ordinal,
                        row: cell.row,
                        column: cell.column,
                    }),
                    _ => None,
                });
                self.current_paragraph = Some(ParagraphBuilder {
                    package_paragraph_id,
                    text: String::new(),
                    utf16_len: 0,
                    formatting: Vec::new(),
                    structure,
                    properties: ParagraphProperties::default(),
                    paragraph_mark_revision: None,
                });
                Frame::Paragraph
            }
            b"r" if self.current_paragraph.is_some() => Frame::Run(RunFrame {
                text: String::new(),
                styles: DirectTextStyles::default(),
                hidden: false,
                direct_child_count: 0,
            }),
            b"rPr" => {
                let Some((run_frame, child_index)) = direct_run_child else {
                    self.frames.push(Frame::RunProperties(RunPropertiesFrame {
                        run_frame: usize::MAX,
                        hidden_eligible: false,
                    }));
                    return Ok(());
                };
                Frame::RunProperties(RunPropertiesFrame {
                    run_frame,
                    hidden_eligible: child_index == 0,
                })
            }
            b"pPr" => Frame::ParagraphProperties,
            b"trPr" => Frame::TableRowProperties,
            b"tcPr" => Frame::TableCellProperties,
            b"pStyle" if matches!(self.frames.last(), Some(Frame::ParagraphProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.style_id = attribute(reader, element, b"val")?;
                }
                Frame::Other
            }
            b"ind" if matches!(self.frames.last(), Some(Frame::ParagraphProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.indentation = parse_indentation(reader, element)?;
                }
                Frame::Other
            }
            b"outlineLvl" if matches!(self.frames.last(), Some(Frame::ParagraphProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.outline_level =
                        Some(parse_outline_level_attribute(reader, element)?);
                }
                Frame::Other
            }
            b"numPr" if matches!(self.frames.last(), Some(Frame::ParagraphProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.numbering.present = true;
                }
                Frame::NumberingProperties
            }
            b"numId" if matches!(self.frames.last(), Some(Frame::NumberingProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.numbering.num_id =
                        Some(parse_u32_attribute(reader, element)?);
                }
                Frame::Other
            }
            b"ilvl" if matches!(self.frames.last(), Some(Frame::NumberingProperties)) => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.properties.numbering.level =
                        Some(parse_level_attribute(reader, element)?);
                }
                Frame::Other
            }
            b"bookmarkStart" => {
                self.start_bookmark(reader, element)?;
                Frame::Other
            }
            b"bookmarkEnd" => {
                self.end_bookmark(reader, element)?;
                Frame::Other
            }
            b"hyperlink" => {
                let anchor = attribute(reader, element, b"anchor")?;
                let reference =
                    anchor
                        .clone()
                        .filter(|anchor| !anchor.is_empty())
                        .and_then(|anchor| {
                            self.current_point().map(|source| RawInternalReference {
                                reference_id: anchor,
                                source,
                            })
                        });
                if anchor.is_some() && reference.is_none() {
                    self.references_complete = false;
                }
                Frame::Hyperlink(reference)
            }
            b"fldSimple" => {
                self.simple_field(reader, element)?;
                Frame::Other
            }
            b"fldChar" => {
                self.field_character(reader, element)?;
                Frame::Other
            }
            b"instrText" | b"delInstrText" => Frame::PseudoText(PseudoTextFrame::instruction()),
            b"ins" | b"del" | b"moveFrom" | b"moveTo" => {
                if self.inside_table_row_properties() {
                    self.revision_unsupported
                        .insert(RevisionUnsupportedReason::StructuralTableRevision);
                    Frame::Other
                } else if self.inside_paragraph_mark_properties() && matches!(name, b"ins" | b"del")
                {
                    if let Some(paragraph) = self.current_paragraph.as_mut() {
                        paragraph.paragraph_mark_revision = Some(if name == b"ins" {
                            ParagraphMarkRevision::Insertion
                        } else {
                            ParagraphMarkRevision::Deletion
                        });
                    }
                    Frame::Other
                } else {
                    Frame::Revision(RevisionFrame {
                        review_index: attributed_revision,
                        start: self.current_review_point(),
                        suppressed: revision_is_suppressed(name, self.revision_view),
                        text: String::new(),
                    })
                }
            }
            b"cellIns" | b"cellDel" | b"cellMerge" => {
                self.revision_unsupported
                    .insert(RevisionUnsupportedReason::StructuralTableRevision);
                Frame::Other
            }
            b"sdt" => {
                let (text_utf8_start, text_utf16_start) =
                    self.current_paragraph.as_ref().map_or((0, 0), |paragraph| {
                        (paragraph.text.len(), paragraph.utf16_len)
                    });
                Frame::Sdt(SdtFrame {
                    placeholder: false,
                    text_utf8_start,
                    text_utf16_start,
                })
            }
            b"showingPlcHdr" => {
                if on_off_element_enabled(reader, element)? {
                    self.mark_placeholder();
                }
                Frame::Other
            }
            b"vanish" => {
                if on_off_element_enabled(reader, element)? {
                    self.mark_hidden_run();
                }
                Frame::Other
            }
            b"rStyle" => {
                if matches!(
                    attribute(reader, element, b"val")?.as_deref(),
                    Some("HideTWBInt" | "HideTWBExt")
                ) {
                    self.mark_hidden_run();
                }
                Frame::Other
            }
            b"b" | b"bCs" => {
                if on_off_element_enabled(reader, element)? {
                    self.mark_bold_run();
                }
                Frame::Other
            }
            b"highlight" => {
                if attribute(reader, element, b"val")?
                    .as_deref()
                    .is_some_and(is_semantic_highlight_color)
                {
                    self.mark_highlighted_run();
                }
                Frame::Other
            }
            b"vertAlign" => {
                if attribute(reader, element, b"val")?.as_deref() == Some("superscript") {
                    self.mark_superscript_run();
                }
                Frame::Other
            }
            b"t" | b"delText" => {
                Frame::PseudoText(PseudoTextFrame::text(preserves_xml_space(reader, element)?))
            }
            b"tab" | b"ptab" => {
                self.append_text_control(TextControl::Tab)?;
                Frame::Other
            }
            b"br" => {
                let control = match attribute(reader, element, b"type")?.as_deref() {
                    Some("page") => TextControl::PageBreak,
                    Some("column") => TextControl::ColumnBreak,
                    _ => TextControl::LineBreak,
                };
                self.append_text_control(control)?;
                Frame::Other
            }
            b"cr" => {
                self.append_text_control(TextControl::CarriageReturn)?;
                Frame::Other
            }
            b"footnoteReference" => {
                let visible =
                    attribute(reader, element, b"customMarkFollows")?.is_none_or(|value| {
                        matches!(value.to_ascii_lowercase().as_str(), "0" | "false" | "off")
                    });
                if visible {
                    self.append_text_control(TextControl::FootnoteReference)?;
                }
                Frame::Other
            }
            b"softHyphen" => {
                self.append_text_control(TextControl::SoftHyphen)?;
                Frame::Other
            }
            b"noBreakHyphen" => {
                self.append_text_control(TextControl::NoBreakHyphen)?;
                Frame::Other
            }
            b"sym" => {
                if let Some(value) = attribute(reader, element, b"char")? {
                    self.append_pseudo_text(&symbol_text(&value))?;
                }
                Frame::Other
            }
            _ => Frame::Other,
        };
        self.frames.push(frame);
        Ok(())
    }

    fn end(&mut self) -> Result<(), ProjectionError> {
        let frame = self
            .frames
            .pop()
            .ok_or(ProjectionError::InvalidDocumentXml)?;
        match frame {
            Frame::Paragraph => {
                let paragraph = self
                    .current_paragraph
                    .take()
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                let ordinal = self.paragraphs.len();
                if let Some(revision) = paragraph.paragraph_mark_revision {
                    self.paragraph_mark_revisions.insert(ordinal, revision);
                }
                self.paragraphs.push(RawProjectedParagraph {
                    ordinal: self.paragraphs.len(),
                    package_paragraph_id: paragraph.package_paragraph_id,
                    text: paragraph.text,
                    utf16_len: paragraph.utf16_len,
                    formatting: paragraph.formatting,
                    structure: paragraph.structure,
                    properties: paragraph.properties,
                });
            }
            Frame::Hyperlink(Some(reference)) => self.references.push(reference),
            Frame::PseudoText(frame) => match frame.kind {
                PseudoTextKind::Text { preserve_space } => {
                    let text = if preserve_space {
                        frame.text.as_str()
                    } else {
                        trim_xml_whitespace(&frame.text)
                    };
                    self.append_pseudo_text(text)?;
                }
                PseudoTextKind::Instruction => self.append_field_instruction(&frame.text),
            },
            Frame::Run(run) if !run.hidden && !self.pseudo_text_is_suppressed() => {
                if let Some(paragraph) = self.current_paragraph.as_mut() {
                    paragraph.append(&run.text, run.styles)?;
                }
            }
            Frame::Revision(revision) => self.finish_attributed_revision(revision)?,
            _ => {}
        }
        Ok(())
    }

    fn record_attributed_revision(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        name: &[u8],
    ) -> Result<Option<usize>, ProjectionError> {
        let Some(kind) = revision_fact_kind(name) else {
            return Ok(None);
        };
        let ReviewRevisionCollection::Complete {
            maximum_facts,
            revisions,
            ..
        } = &mut self.review_revisions
        else {
            return Ok(None);
        };
        if revisions.len() >= *maximum_facts {
            self.review_revisions = ReviewRevisionCollection::LimitExceeded;
            return Ok(None);
        }
        let review_index = revisions.len();
        revisions.push(AttributedRevision {
            kind,
            author: attribute(reader, element, b"author")?.unwrap_or_default(),
            date: attribute(reader, element, b"date")?,
            revision_id: attribute(reader, element, b"id")?,
            content: ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation),
        });
        Ok(Some(review_index))
    }

    fn finish_attributed_revision(
        &mut self,
        revision: RevisionFrame,
    ) -> Result<(), ProjectionError> {
        let (Some(review_index), Some(start), Some(end)) = (
            revision.review_index,
            revision.start,
            self.current_review_point(),
        ) else {
            return Ok(());
        };
        let ReviewRevisionCollection::Complete { revisions, .. } = &mut self.review_revisions
        else {
            return Ok(());
        };
        let attributed = revisions
            .get_mut(review_index)
            .ok_or(ProjectionError::InvalidDocumentXml)?;
        attributed.content = ReviewDetail::Known(RevisionContent {
            span: ReviewSpan { start, end },
            formatting_only: revision.text.is_empty(),
            text: revision.text,
        });
        Ok(())
    }

    fn inside_body(&self) -> bool {
        self.frames.iter().any(|frame| matches!(frame, Frame::Body))
    }

    fn inside_textbox(&self) -> bool {
        self.frames
            .iter()
            .any(|frame| matches!(frame, Frame::Textbox))
    }

    fn inside_change_snapshot(&self) -> bool {
        self.frames
            .iter()
            .any(|frame| matches!(frame, Frame::ChangeSnapshot))
    }

    fn inside_paragraph_mark_properties(&self) -> bool {
        self.current_paragraph.is_some()
            && self
                .frames
                .iter()
                .any(|frame| matches!(frame, Frame::ParagraphProperties))
            && !self
                .frames
                .iter()
                .any(|frame| matches!(frame, Frame::Run(_)))
    }

    fn inside_table_row_properties(&self) -> bool {
        self.frames
            .iter()
            .any(|frame| matches!(frame, Frame::TableRowProperties))
    }

    fn pseudo_text_is_suppressed(&self) -> bool {
        self.frames.iter().any(|frame| {
            matches!(
                frame,
                Frame::ParagraphProperties
                    | Frame::RunProperties(_)
                    | Frame::ChangeSnapshot
                    | Frame::Textbox
            ) || matches!(frame, Frame::Revision(revision) if revision.suppressed)
                || matches!(frame, Frame::Sdt(sdt) if sdt.placeholder)
        })
    }

    fn pseudo_text_is_structurally_suppressed(&self) -> bool {
        self.frames.iter().any(|frame| {
            matches!(
                frame,
                Frame::ParagraphProperties
                    | Frame::RunProperties(_)
                    | Frame::ChangeSnapshot
                    | Frame::Textbox
            ) || matches!(frame, Frame::Sdt(sdt) if sdt.placeholder)
        })
    }

    fn append_pseudo_text(&mut self, text: &str) -> Result<(), ProjectionError> {
        if self.current_paragraph.is_none()
            || self.pseudo_text_is_structurally_suppressed()
            || self
                .frames
                .iter()
                .rev()
                .any(|frame| matches!(frame, Frame::Run(run) if run.hidden))
        {
            return Ok(());
        }
        self.append_review_text(text);
        if self.pseudo_text_is_suppressed() {
            return Ok(());
        }
        if let Some(run) = self.frames.iter_mut().rev().find_map(|frame| match frame {
            Frame::Run(run) => Some(run),
            _ => None,
        }) {
            run.text.push_str(text);
            return Ok(());
        }
        self.current_paragraph
            .as_mut()
            .ok_or(ProjectionError::InvalidDocumentXml)?
            .append(text, DirectTextStyles::default())
    }

    fn append_review_text(&mut self, text: &str) {
        let ReviewRevisionCollection::Complete {
            remaining_detail_bytes,
            ..
        } = &mut self.review_revisions
        else {
            return;
        };
        let mut limit_exceeded = false;
        for frame in &mut self.frames {
            if let Frame::Revision(revision) = frame
                && revision.review_index.is_some()
            {
                let Some(remaining) = remaining_detail_bytes.checked_sub(text.len()) else {
                    limit_exceeded = true;
                    break;
                };
                *remaining_detail_bytes = remaining;
                revision.text.push_str(text);
            }
        }
        if limit_exceeded {
            self.review_revisions = ReviewRevisionCollection::LimitExceeded;
        }
    }

    fn current_review_point(&self) -> Option<ReviewPoint> {
        let paragraph = self.current_paragraph.as_ref()?;
        let run_text = self.frames.iter().rev().find_map(|frame| match frame {
            Frame::Run(run) if !run.hidden && !self.pseudo_text_is_suppressed() => {
                Some(run.text.as_str())
            }
            _ => None,
        });
        let run_utf8 = run_text.map_or(0, str::len);
        let run_utf16 = run_text.map_or(0, |text| text.encode_utf16().count());
        Some(ReviewPoint {
            paragraph_ordinal: self.paragraphs.len(),
            utf8: u32::try_from(paragraph.text.len().checked_add(run_utf8)?).ok()?,
            utf16: paragraph
                .utf16_len
                .checked_add(u32::try_from(run_utf16).ok()?)?,
        })
    }

    fn next_paragraph_start_review_point(&self) -> Option<ReviewPoint> {
        if self.current_paragraph.is_some() {
            return self.current_review_point();
        }
        self.inside_body().then_some(ReviewPoint {
            paragraph_ordinal: self.paragraphs.len(),
            utf8: 0,
            utf16: 0,
        })
    }

    fn previous_paragraph_end_review_point(&self) -> Option<ReviewPoint> {
        if self.current_paragraph.is_some() {
            return self.current_review_point();
        }
        let paragraph = self.paragraphs.last()?;
        Some(ReviewPoint {
            paragraph_ordinal: paragraph.ordinal,
            utf8: u32::try_from(paragraph.text.len()).ok()?,
            utf16: paragraph.utf16_len,
        })
    }

    fn review_comment_id(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Option<String>, ProjectionError> {
        attribute(reader, element, b"id")
    }

    fn start_review_comment_anchor(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let (Some(comment_id), Some(point)) = (
            Self::review_comment_id(reader, element)?,
            self.next_paragraph_start_review_point(),
        ) else {
            return Ok(());
        };
        if self.invalid_review_comment_anchors.contains(&comment_id) {
            return Ok(());
        }
        if self
            .open_review_comment_anchors
            .insert(comment_id.clone(), point)
            .is_some()
        {
            self.open_review_comment_anchors.remove(&comment_id);
            self.review_comment_anchors.remove(&comment_id);
            self.invalid_review_comment_anchors.insert(comment_id);
        }
        Ok(())
    }

    fn suppress_review_comment_anchor(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let Some(comment_id) = Self::review_comment_id(reader, element)? else {
            return Ok(());
        };
        self.open_review_comment_anchors.remove(&comment_id);
        self.review_comment_anchors.remove(&comment_id);
        self.invalid_review_comment_anchors.insert(comment_id);
        Ok(())
    }

    fn end_review_comment_anchor(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let (Some(comment_id), Some(end)) = (
            Self::review_comment_id(reader, element)?,
            self.previous_paragraph_end_review_point(),
        ) else {
            return Ok(());
        };
        if self.invalid_review_comment_anchors.contains(&comment_id) {
            return Ok(());
        }
        let Some(start) = self.open_review_comment_anchors.remove(&comment_id) else {
            self.review_comment_anchors.remove(&comment_id);
            self.invalid_review_comment_anchors.insert(comment_id);
            return Ok(());
        };
        if self
            .review_comment_anchors
            .insert(comment_id.clone(), ReviewSpan { start, end })
            .is_some()
        {
            self.review_comment_anchors.remove(&comment_id);
            self.invalid_review_comment_anchors.insert(comment_id);
        }
        Ok(())
    }

    fn record_review_comment_reference(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let (Some(comment_id), Some(point)) = (
            Self::review_comment_id(reader, element)?,
            self.current_review_point(),
        ) else {
            return Ok(());
        };
        if self.invalid_review_comment_anchors.contains(&comment_id) {
            return Ok(());
        }
        if !self.review_comment_anchors.contains_key(&comment_id)
            && !self.open_review_comment_anchors.contains_key(&comment_id)
        {
            self.review_comment_anchors.insert(
                comment_id,
                ReviewSpan {
                    start: point,
                    end: point,
                },
            );
        }
        Ok(())
    }

    fn append_pseudo_content(&mut self, text: &str) -> Result<(), ProjectionError> {
        let Some(Frame::PseudoText(frame)) = self.frames.last_mut() else {
            return Err(ProjectionError::InvalidDocumentXml);
        };
        frame.text.push_str(text);
        Ok(())
    }

    fn append_text_control(&mut self, control: TextControl) -> Result<(), ProjectionError> {
        if let Some(text) = control.materialize(self.text_materialization) {
            self.append_pseudo_text(text)?;
        }
        Ok(())
    }

    fn append_field_instruction(&mut self, text: &str) {
        let Some(field) = self.fields.last_mut() else {
            self.references_complete = false;
            return;
        };
        if field.separated || field.instruction.len().saturating_add(text.len()) > 4096 {
            self.references_complete = false;
            return;
        }
        field.instruction.push_str(text);
    }

    fn simple_field(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let instruction = attribute(reader, element, b"instr")?;
        let source = self.current_point();
        let Some(instruction) = instruction.filter(|value| value.len() <= 4096) else {
            self.references_complete = false;
            return Ok(());
        };
        self.record_field_reference(&instruction, source);
        Ok(())
    }

    fn field_character(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        match attribute(reader, element, b"fldCharType")?.as_deref() {
            Some("begin") => self.fields.push(FieldFrame {
                source: self.current_field_point(),
                instruction: String::new(),
                separated: false,
            }),
            Some("separate") => {
                let Some(mut field) = self.fields.pop() else {
                    self.references_complete = false;
                    return Ok(());
                };
                if field.separated {
                    self.references_complete = false;
                } else {
                    self.record_field_reference(&field.instruction, field.source.clone());
                    field.separated = true;
                }
                self.fields.push(field);
            }
            Some("end") => {
                let Some(field) = self.fields.pop() else {
                    self.references_complete = false;
                    return Ok(());
                };
                if !field.separated {
                    self.record_field_reference(&field.instruction, field.source);
                }
            }
            _ => self.references_complete = false,
        }
        Ok(())
    }

    fn current_field_point(&self) -> Option<RawBlockPoint> {
        let paragraph = self.current_paragraph.as_ref()?;
        if self.pseudo_text_is_suppressed() {
            return None;
        }
        if self.frames.iter().rev().find_map(|frame| match frame {
            Frame::Run(run) => Some(!run.text.is_empty()),
            _ => None,
        }) == Some(true)
        {
            return None;
        }
        Some(RawBlockPoint {
            paragraph: self.paragraphs.len(),
            utf8: u32::try_from(paragraph.text.len()).ok()?,
            utf16: paragraph.utf16_len,
        })
    }

    fn record_field_reference(&mut self, instruction: &str, source: Option<RawBlockPoint>) {
        match internal_reference_target(instruction) {
            Ok(Some(reference_id)) => {
                let Some(source) = source else {
                    self.references_complete = false;
                    return;
                };
                self.references.push(RawInternalReference {
                    reference_id,
                    source,
                });
            }
            Ok(None) => {}
            Err(()) => self.references_complete = false,
        }
    }

    fn current_point(&self) -> Option<RawBlockPoint> {
        if let Some(paragraph) = self.current_paragraph.as_ref() {
            if self.pseudo_text_is_suppressed()
                || self
                    .frames
                    .iter()
                    .any(|frame| matches!(frame, Frame::Run(_)))
            {
                return None;
            }
            return Some(RawBlockPoint {
                paragraph: self.paragraphs.len(),
                utf8: u32::try_from(paragraph.text.len()).ok()?,
                utf16: paragraph.utf16_len,
            });
        }
        if !matches!(self.frames.last(), Some(Frame::Body)) {
            return None;
        }
        self.paragraphs.last().map_or(
            Some(RawBlockPoint {
                paragraph: 0,
                utf8: 0,
                utf16: 0,
            }),
            |paragraph| {
                Some(RawBlockPoint {
                    paragraph: paragraph.ordinal,
                    utf8: u32::try_from(paragraph.text.len()).ok()?,
                    utf16: u32::try_from(paragraph.text.encode_utf16().count()).ok()?,
                })
            },
        )
    }

    fn start_bookmark(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let id = attribute(reader, element, b"id")?.and_then(|value| value.parse::<u32>().ok());
        let name = attribute(reader, element, b"name")?;
        let point = self.current_point();
        let (Some(id), Some(name), Some(point)) = (id, name, point) else {
            self.bookmarks_complete = false;
            return Ok(());
        };
        if name.is_empty()
            || !self.seen_bookmark_ids.insert(id)
            || self.open_bookmarks.insert(id, (name, point)).is_some()
        {
            self.bookmarks_complete = false;
        }
        Ok(())
    }

    fn end_bookmark(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ProjectionError> {
        let id = attribute(reader, element, b"id")?.and_then(|value| value.parse::<u32>().ok());
        let point = self.current_point();
        let (Some(id), Some(end)) = (id, point) else {
            self.bookmarks_complete = false;
            return Ok(());
        };
        let Some((name, start)) = self.open_bookmarks.remove(&id) else {
            self.bookmarks_complete = false;
            return Ok(());
        };
        self.bookmarks.push(RawBookmarkRange {
            id,
            name,
            start,
            end,
        });
        Ok(())
    }

    fn mark_placeholder(&mut self) {
        let Some(sdt) = self.frames.iter_mut().rev().find_map(|frame| match frame {
            Frame::Sdt(sdt) => Some(sdt),
            _ => None,
        }) else {
            return;
        };
        sdt.placeholder = true;
        if let Some(paragraph) = self.current_paragraph.as_mut() {
            paragraph.truncate(sdt.text_utf8_start, sdt.text_utf16_start);
        }
    }

    fn mark_hidden_run(&mut self) {
        let Some((run_frame, hidden_eligible)) =
            self.frames.iter().rev().find_map(|frame| match frame {
                Frame::RunProperties(properties) => {
                    Some((properties.run_frame, properties.hidden_eligible))
                }
                _ => None,
            })
        else {
            return;
        };
        if !hidden_eligible || run_frame == usize::MAX {
            return;
        }
        if let Some(Frame::Run(run)) = self.frames.get_mut(run_frame) {
            run.hidden = true;
        }
    }

    fn mark_bold_run(&mut self) {
        let Some(Frame::RunProperties(properties)) = self.frames.last() else {
            return;
        };
        let run_frame = properties.run_frame;
        if let Some(Frame::Run(run)) = self.frames.get_mut(run_frame) {
            run.styles.bold = true;
        }
    }

    fn mark_highlighted_run(&mut self) {
        let Some(Frame::RunProperties(properties)) = self.frames.last() else {
            return;
        };
        let run_frame = properties.run_frame;
        if let Some(Frame::Run(run)) = self.frames.get_mut(run_frame) {
            run.styles.highlighted = true;
        }
    }

    fn mark_superscript_run(&mut self) {
        let Some(Frame::RunProperties(properties)) = self.frames.last() else {
            return;
        };
        let run_frame = properties.run_frame;
        if let Some(Frame::Run(run)) = self.frames.get_mut(run_frame) {
            run.styles.superscript = true;
        }
    }
}

const fn revision_fact_kind(name: &[u8]) -> Option<RevisionFactKind> {
    match name {
        b"ins" => Some(RevisionFactKind::Insertion),
        b"del" => Some(RevisionFactKind::Deletion),
        b"moveFrom" => Some(RevisionFactKind::MoveFrom),
        b"moveTo" => Some(RevisionFactKind::MoveTo),
        b"cellIns" => Some(RevisionFactKind::CellInsertion),
        b"cellDel" => Some(RevisionFactKind::CellDeletion),
        b"cellMerge" => Some(RevisionFactKind::CellMerge),
        b"pPrChange" => Some(RevisionFactKind::ParagraphPropertiesChange),
        b"rPrChange" => Some(RevisionFactKind::RunPropertiesChange),
        b"sectPrChange" => Some(RevisionFactKind::SectionPropertiesChange),
        b"tblPrChange" => Some(RevisionFactKind::TablePropertiesChange),
        b"trPrChange" => Some(RevisionFactKind::TableRowPropertiesChange),
        b"tcPrChange" => Some(RevisionFactKind::TableCellPropertiesChange),
        b"tblGridChange" => Some(RevisionFactKind::TableGridChange),
        b"customXmlDelRangeStart" => Some(RevisionFactKind::CustomXmlDeletionRangeStart),
        b"customXmlDelRangeEnd" => Some(RevisionFactKind::CustomXmlDeletionRangeEnd),
        b"customXmlInsRangeStart" => Some(RevisionFactKind::CustomXmlInsertionRangeStart),
        b"customXmlInsRangeEnd" => Some(RevisionFactKind::CustomXmlInsertionRangeEnd),
        b"customXmlMoveFromRangeStart" => Some(RevisionFactKind::CustomXmlMoveFromRangeStart),
        b"customXmlMoveFromRangeEnd" => Some(RevisionFactKind::CustomXmlMoveFromRangeEnd),
        b"customXmlMoveToRangeStart" => Some(RevisionFactKind::CustomXmlMoveToRangeStart),
        b"customXmlMoveToRangeEnd" => Some(RevisionFactKind::CustomXmlMoveToRangeEnd),
        _ => None,
    }
}

fn enclosing_table(frames: &mut [Frame]) -> Option<&mut TableFrame> {
    for frame in frames.iter_mut().rev() {
        match frame {
            Frame::Table(table) => return Some(table),
            Frame::Row(_) | Frame::Cell(_) | Frame::Paragraph | Frame::Run(_) => return None,
            _ => {}
        }
    }
    None
}

fn enclosing_row(frames: &mut [Frame]) -> Option<&mut RowFrame> {
    for frame in frames.iter_mut().rev() {
        match frame {
            Frame::Row(row) => return Some(row),
            Frame::Table(_) | Frame::Cell(_) | Frame::Paragraph | Frame::Run(_) => return None,
            _ => {}
        }
    }
    None
}

fn revision_is_suppressed(name: &[u8], view: RevisionView) -> bool {
    match view {
        RevisionView::Current => matches!(name, b"del" | b"moveFrom"),
        RevisionView::Original => matches!(name, b"ins" | b"moveTo"),
    }
}

fn is_change_snapshot(name: &[u8]) -> bool {
    matches!(
        name,
        b"pPrChange"
            | b"rPrChange"
            | b"sectPrChange"
            | b"tblPrChange"
            | b"trPrChange"
            | b"tcPrChange"
            | b"tblGridChange"
    )
}

fn is_unsupported_revision_markup(name: &[u8]) -> bool {
    matches!(
        name,
        b"customXmlDelRangeStart"
            | b"customXmlDelRangeEnd"
            | b"customXmlInsRangeStart"
            | b"customXmlInsRangeEnd"
            | b"customXmlMoveFromRangeStart"
            | b"customXmlMoveFromRangeEnd"
            | b"customXmlMoveToRangeStart"
            | b"customXmlMoveToRangeEnd"
    )
}

const fn paragraph_break_is_removed(
    revision: Option<ParagraphMarkRevision>,
    view: RevisionView,
) -> bool {
    matches!(
        (revision, view),
        (Some(ParagraphMarkRevision::Deletion), RevisionView::Current)
            | (
                Some(ParagraphMarkRevision::Insertion),
                RevisionView::Original
            )
    )
}

fn normalize_paragraph_revision_view(
    paragraphs: &mut Vec<RawProjectedParagraph>,
    paragraph_mark_revisions: &mut HashMap<usize, ParagraphMarkRevision>,
    view: RevisionView,
    unsupported: &mut BTreeSet<RevisionUnsupportedReason>,
) -> Result<bool, ProjectionError> {
    let mut normalized: Vec<RawProjectedParagraph> = Vec::with_capacity(paragraphs.len());
    let mut merge_previous = false;
    let mut merged_any = false;

    for (source_ordinal, mut paragraph) in std::mem::take(paragraphs).into_iter().enumerate() {
        let merge_next =
            paragraph_break_is_removed(paragraph_mark_revisions.remove(&source_ordinal), view);
        if merge_previous {
            let previous = normalized
                .last_mut()
                .ok_or(ProjectionError::InvalidDocumentXml)?;
            if previous.structure != paragraph.structure {
                unsupported.insert(RevisionUnsupportedReason::IncompatibleParagraphMerge);
                paragraph.ordinal = normalized.len();
                normalized.push(paragraph);
                merge_previous = merge_next;
                continue;
            }
            let utf16_offset = previous.utf16_len;
            for mut span in paragraph.formatting {
                span.start_utf16 = span
                    .start_utf16
                    .checked_add(utf16_offset)
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                span.end_utf16 = span
                    .end_utf16
                    .checked_add(utf16_offset)
                    .ok_or(ProjectionError::InvalidDocumentXml)?;
                previous.formatting.push(span);
            }
            previous.text.push_str(&paragraph.text);
            previous.utf16_len = previous
                .utf16_len
                .checked_add(paragraph.utf16_len)
                .ok_or(ProjectionError::InvalidDocumentXml)?;
            merge_previous = merge_next;
            merged_any = true;
            continue;
        }

        paragraph.ordinal = normalized.len();
        normalized.push(paragraph);
        merge_previous = merge_next;
    }

    *paragraphs = normalized;
    Ok(merged_any)
}

fn internal_reference_target(instruction: &str) -> Result<Option<String>, ()> {
    let tokens = field_tokens(instruction)?;
    let Some(command) = tokens.first() else {
        return Err(());
    };
    if ["REF", "PAGEREF", "NOTEREF"]
        .iter()
        .any(|candidate| command.eq_ignore_ascii_case(candidate))
    {
        return tokens
            .get(1)
            .filter(|target| !target.is_empty() && !target.starts_with('\\'))
            .cloned()
            .map(Some)
            .ok_or(());
    }
    if command.eq_ignore_ascii_case("HYPERLINK") {
        for (index, token) in tokens.iter().enumerate().skip(1) {
            if token.eq_ignore_ascii_case("\\l") {
                return tokens
                    .get(index.checked_add(1).ok_or(())?)
                    .filter(|target| !target.is_empty())
                    .cloned()
                    .map(Some)
                    .ok_or(());
            }
        }
    }
    Ok(None)
}

fn field_tokens(instruction: &str) -> Result<Vec<String>, ()> {
    let mut tokens = Vec::new();
    let mut token = String::new();
    let mut quoted = false;
    for character in instruction.chars() {
        match character {
            '"' => quoted = !quoted,
            whitespace if whitespace.is_whitespace() && !quoted => {
                if !token.is_empty() {
                    tokens.push(std::mem::take(&mut token));
                }
            }
            _ => token.push(character),
        }
    }
    if quoted {
        return Err(());
    }
    if !token.is_empty() {
        tokens.push(token);
    }
    Ok(tokens)
}

/// Returns whether an OOXML `ST_HighlightColor` value represents a visible
/// semantic marker. Run shading (`w:shd`) is a distinct formatting fact.
#[must_use]
pub fn is_semantic_highlight_color(value: &str) -> bool {
    matches!(
        value,
        "black"
            | "blue"
            | "cyan"
            | "darkBlue"
            | "darkCyan"
            | "darkGreen"
            | "darkMagenta"
            | "darkRed"
            | "darkYellow"
            | "green"
            | "magenta"
            | "red"
            | "yellow"
    )
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, ProjectionError> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if OoxmlNamespace::from_resolved(&namespace) == OoxmlNamespace::Wordprocessing
            && local_name.as_ref() == name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ProjectionError::InvalidDocumentXml);
        }
    }
    Ok(None)
}

fn preserves_xml_space(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<bool, ProjectionError> {
    const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";

    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(
            namespace,
            quick_xml::name::ResolveResult::Bound(namespace)
                if namespace.as_ref() == XML_NAMESPACE
        ) && local_name.as_ref() == b"space"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|_| ProjectionError::InvalidDocumentXml)?;
            return match value.as_ref() {
                "default" => Ok(false),
                "preserve" => Ok(true),
                _ => Err(ProjectionError::InvalidDocumentXml),
            };
        }
    }
    Ok(false)
}

fn on_off_element_enabled(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<bool, ProjectionError> {
    Ok(attribute(reader, element, b"val")?
        .is_none_or(|value| !matches!(value.to_ascii_lowercase().as_str(), "0" | "false" | "off")))
}

fn trim_xml_whitespace(value: &str) -> &str {
    value.trim_matches(['\u{0009}', '\u{000a}', '\u{000d}', '\u{0020}'])
}

fn attribute_2010(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, ProjectionError> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ProjectionError::InvalidDocumentXml)?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if OoxmlNamespace::from_resolved(&namespace) == OoxmlNamespace::Wordprocessing2010
            && local_name.as_ref() == name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ProjectionError::InvalidDocumentXml);
        }
    }
    Ok(None)
}

fn symbol_text(value: &str) -> String {
    let prefix = value
        .chars()
        .take_while(char::is_ascii_hexdigit)
        .collect::<String>();
    let Some(code) = (!prefix.is_empty())
        .then(|| u32::from_str_radix(&prefix, 16).ok())
        .flatten()
    else {
        return value.to_owned();
    };
    char::from_u32(code).unwrap_or('\u{fffd}').to_string()
}
