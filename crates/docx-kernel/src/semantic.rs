#![forbid(unsafe_code)]

use std::collections::HashMap;

use quick_xml::{
    XmlVersion,
    escape::resolve_predefined_entity,
    events::{BytesStart, Event},
    name::{Namespace, PrefixDeclaration, ResolveResult},
    reader::NsReader,
};
use thiserror::Error;

const WORDPROCESSING_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORDPROCESSING_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const OFFICE_RELATIONSHIPS_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE_RELATIONSHIPS_STRICT: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MARKUP_COMPATIBILITY_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MARKUP_COMPATIBILITY_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/markup-compatibility/main";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ScanLimits {
    pub blocks: usize,
    pub segments: usize,
    pub depth: usize,
    pub events: usize,
    pub inline_context_copies: usize,
}

impl Default for ScanLimits {
    fn default() -> Self {
        Self {
            blocks: 100_000,
            segments: 1_000_000,
            depth: 256,
            events: 10_000_000,
            inline_context_copies: 20_000_000,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ScanWork {
    pub events: usize,
    pub inline_context_copies: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RevisionKind {
    Deletion,
    Insertion,
    MoveFrom,
    MoveTo,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum InlineContext {
    Hyperlink {
        relationship_id: Option<String>,
        anchor: Option<String>,
    },
    Revision(RevisionKind),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SegmentSource {
    Break,
    Tab,
    Text,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextSegment {
    pub start_utf16: usize,
    pub end_utf16: usize,
    pub source: SegmentSource,
    pub contexts: Vec<InlineContext>,
    pub xml_path: Vec<usize>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum BlockLocation {
    Paragraph {
        xml_path: Vec<usize>,
    },
    TableCellParagraph {
        xml_path: Vec<usize>,
        table_path: Vec<usize>,
        row_path: Vec<usize>,
        cell_path: Vec<usize>,
    },
    TextBoxParagraph {
        xml_path: Vec<usize>,
        text_box_path: Vec<usize>,
    },
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextBlock {
    pub text: String,
    pub location: BlockLocation,
    pub segments: Vec<TextSegment>,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct PartCoverage {
    pub hyperlink_segments: usize,
    pub revision_segments: usize,
    pub unsupported_alternate_content: usize,
    pub unsupported_symbols: usize,
    pub unsupported_field_instructions: usize,
    pub work: ScanWork,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PartScan {
    pub blocks: Vec<TextBlock>,
    pub coverage: PartCoverage,
}

#[derive(Clone, Copy, Debug, Error, Eq, PartialEq)]
pub enum ScanError {
    #[error("WordprocessingML part is invalid XML")]
    InvalidXml,
    #[error("WordprocessingML part must not contain a document type declaration")]
    DocumentTypeDeclaration,
    #[error("WordprocessingML text is outside a paragraph")]
    TextOutsideParagraph,
    #[error("WordprocessingML part exceeds the configured nesting depth")]
    TooDeep,
    #[error("WordprocessingML part exceeds the configured block limit")]
    TooManyBlocks,
    #[error("WordprocessingML part exceeds the configured segment limit")]
    TooManySegments,
    #[error("WordprocessingML part exceeds the configured XML event limit")]
    TooManyEvents,
    #[error("WordprocessingML part exceeds the configured inline-context copy limit")]
    TooManyInlineContextCopies,
    #[error("WordprocessingML text length overflowed")]
    TextLengthOverflow,
    #[error("WordprocessingML part must be UTF-8 encoded")]
    UnsupportedEncoding,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum NamespaceKind {
    MarkupCompatibility,
    OfficeRelationships,
    Other,
    Wordprocessing,
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    let ResolveResult::Bound(Namespace(value)) = namespace else {
        return NamespaceKind::Other;
    };
    if matches!(*value, WORDPROCESSING_TRANSITIONAL | WORDPROCESSING_STRICT) {
        NamespaceKind::Wordprocessing
    } else if matches!(
        *value,
        OFFICE_RELATIONSHIPS_TRANSITIONAL | OFFICE_RELATIONSHIPS_STRICT
    ) {
        NamespaceKind::OfficeRelationships
    } else if matches!(
        *value,
        MARKUP_COMPATIBILITY_TRANSITIONAL | MARKUP_COMPATIBILITY_STRICT
    ) {
        NamespaceKind::MarkupCompatibility
    } else {
        NamespaceKind::Other
    }
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    local_name: &[u8],
) -> Result<Option<String>, ScanError> {
    for item in element.attributes() {
        let item = item.map_err(|_| ScanError::InvalidXml)?;
        let (resolved, name) = reader.resolver().resolve_attribute(item.key);
        if namespace_kind(&resolved) == namespace && name.as_ref() == local_name {
            return item
                .decoded_and_normalized_value(XmlVersion::default(), reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ScanError::InvalidXml);
        }
    }
    Ok(None)
}

fn unqualified_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>, ScanError> {
    for item in element.attributes() {
        let item = item.map_err(|_| ScanError::InvalidXml)?;
        let (resolved, name) = reader.resolver().resolve_attribute(item.key);
        if matches!(resolved, ResolveResult::Unbound) && name.as_ref() == local_name {
            return item
                .decoded_and_normalized_value(XmlVersion::default(), reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ScanError::InvalidXml);
        }
    }
    Ok(None)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum FrameKind {
    AlternateContent { branch_selected: bool },
    AlternateContentBranch,
    Other,
    Paragraph(usize),
    Run,
    Text,
}

struct Frame {
    kind: FrameKind,
    next_child: usize,
    context_pushed: bool,
    table_path: Option<Vec<usize>>,
    row_path: Option<Vec<usize>>,
    cell_path: Option<Vec<usize>>,
    text_box_path: Option<Vec<usize>>,
    content_suppressed: bool,
}

impl Frame {
    const fn root() -> Self {
        Self {
            kind: FrameKind::Other,
            next_child: 0,
            context_pushed: false,
            table_path: None,
            row_path: None,
            cell_path: None,
            text_box_path: None,
            content_suppressed: false,
        }
    }
}

struct BlockBuilder {
    text: String,
    utf16_len: usize,
    location: BlockLocation,
    segments: Vec<TextSegment>,
}

struct PendingText {
    block_id: Option<usize>,
    path: Vec<usize>,
    value: String,
}

struct Scanner<F> {
    frames: Vec<Frame>,
    element_path: Vec<usize>,
    builders: HashMap<usize, BlockBuilder>,
    completed: HashMap<usize, TextBlock>,
    coverage: PartCoverage,
    active_contexts: Vec<InlineContext>,
    limits: ScanLimits,
    next_block_id: usize,
    next_emit_id: usize,
    segment_count: usize,
    pending_text: Option<PendingText>,
    root_seen: bool,
    sink: F,
}

impl<F> Scanner<F>
where
    F: FnMut(TextBlock) -> Result<(), ScanError>,
{
    fn new(limits: ScanLimits, sink: F) -> Self {
        Self {
            frames: vec![Frame::root()],
            element_path: Vec::new(),
            builders: HashMap::new(),
            completed: HashMap::new(),
            coverage: PartCoverage::default(),
            active_contexts: Vec::new(),
            limits,
            next_block_id: 0,
            next_emit_id: 0,
            segment_count: 0,
            pending_text: None,
            root_seen: false,
            sink,
        }
    }

    fn enter_element(&mut self) -> Result<(), ScanError> {
        let parent = self.frames.last_mut().ok_or(ScanError::InvalidXml)?;
        let child = parent.next_child;
        parent.next_child = parent
            .next_child
            .checked_add(1)
            .ok_or(ScanError::InvalidXml)?;
        self.element_path.push(child);
        Ok(())
    }

    fn current_block_id(&self) -> Option<usize> {
        self.frames.iter().rev().find_map(|frame| match frame.kind {
            FrameKind::Paragraph(block_id) => Some(block_id),
            FrameKind::AlternateContent { .. }
            | FrameKind::AlternateContentBranch
            | FrameKind::Other
            | FrameKind::Run
            | FrameKind::Text => None,
        })
    }

    fn inside_run(&self) -> bool {
        self.frames
            .iter()
            .rev()
            .any(|frame| frame.kind == FrameKind::Run)
    }

    fn contexts(&mut self) -> Result<Vec<InlineContext>, ScanError> {
        let copies = self
            .coverage
            .work
            .inline_context_copies
            .checked_add(self.active_contexts.len())
            .ok_or(ScanError::TooManyInlineContextCopies)?;
        if copies > self.limits.inline_context_copies {
            return Err(ScanError::TooManyInlineContextCopies);
        }
        self.coverage.work.inline_context_copies = copies;
        Ok(self.active_contexts.clone())
    }

    fn record_event(&mut self) -> Result<(), ScanError> {
        let events = self
            .coverage
            .work
            .events
            .checked_add(1)
            .ok_or(ScanError::TooManyEvents)?;
        if events > self.limits.events {
            return Err(ScanError::TooManyEvents);
        }
        self.coverage.work.events = events;
        Ok(())
    }

    fn nearest_path(&self, select: impl Fn(&Frame) -> Option<&Vec<usize>>) -> Option<Vec<usize>> {
        self.frames.iter().rev().find_map(select).cloned()
    }

    fn record_coverage(&mut self, namespace: NamespaceKind, name: &[u8]) -> Result<(), ScanError> {
        if namespace == NamespaceKind::MarkupCompatibility && name == b"AlternateContent" {
            self.coverage.unsupported_alternate_content = self
                .coverage
                .unsupported_alternate_content
                .checked_add(1)
                .ok_or(ScanError::TooManySegments)?;
        } else if namespace == NamespaceKind::Wordprocessing && name == b"sym" {
            self.coverage.unsupported_symbols = self
                .coverage
                .unsupported_symbols
                .checked_add(1)
                .ok_or(ScanError::TooManySegments)?;
        } else if namespace == NamespaceKind::Wordprocessing
            && matches!(name, b"instrText" | b"fldSimple")
        {
            self.coverage.unsupported_field_instructions = self
                .coverage
                .unsupported_field_instructions
                .checked_add(1)
                .ok_or(ScanError::TooManySegments)?;
        }
        Ok(())
    }

    fn block_location(&self, path: &[usize]) -> BlockLocation {
        self.nearest_path(|frame| frame.text_box_path.as_ref())
            .map_or_else(
                || {
                    let table = self.nearest_path(|frame| frame.table_path.as_ref());
                    let row = self.nearest_path(|frame| frame.row_path.as_ref());
                    let cell = self.nearest_path(|frame| frame.cell_path.as_ref());
                    if let (Some(table_path), Some(row_path), Some(cell_path)) = (table, row, cell)
                    {
                        BlockLocation::TableCellParagraph {
                            xml_path: path.to_vec(),
                            table_path,
                            row_path,
                            cell_path,
                        }
                    } else {
                        BlockLocation::Paragraph {
                            xml_path: path.to_vec(),
                        }
                    }
                },
                |text_box| BlockLocation::TextBoxParagraph {
                    xml_path: path.to_vec(),
                    text_box_path: text_box,
                },
            )
    }

    fn begin_paragraph(&mut self, path: &[usize]) -> Result<FrameKind, ScanError> {
        if self.next_block_id >= self.limits.blocks {
            return Err(ScanError::TooManyBlocks);
        }
        let block_id = self.next_block_id;
        self.next_block_id = self
            .next_block_id
            .checked_add(1)
            .ok_or(ScanError::TooManyBlocks)?;
        self.builders.insert(
            block_id,
            BlockBuilder {
                text: String::new(),
                utf16_len: 0,
                location: self.block_location(path),
                segments: Vec::new(),
            },
        );
        Ok(FrameKind::Paragraph(block_id))
    }

    fn inline_context(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        name: &[u8],
    ) -> Result<Option<InlineContext>, ScanError> {
        match name {
            b"hyperlink" => Ok(Some(InlineContext::Hyperlink {
                relationship_id: attribute(
                    reader,
                    element,
                    NamespaceKind::OfficeRelationships,
                    b"id",
                )?,
                anchor: attribute(reader, element, NamespaceKind::Wordprocessing, b"anchor")?,
            })),
            b"del" => Ok(Some(InlineContext::Revision(RevisionKind::Deletion))),
            b"ins" => Ok(Some(InlineContext::Revision(RevisionKind::Insertion))),
            b"moveFrom" => Ok(Some(InlineContext::Revision(RevisionKind::MoveFrom))),
            b"moveTo" => Ok(Some(InlineContext::Revision(RevisionKind::MoveTo))),
            _ => Ok(None),
        }
    }

    fn alternate_choice_supported(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<bool, ScanError> {
        let Some(requires) = unqualified_attribute(reader, element, b"Requires")? else {
            return Ok(false);
        };
        let mut prefixes = requires.split_ascii_whitespace().peekable();
        if prefixes.peek().is_none() {
            return Ok(false);
        }
        Ok(prefixes.all(|required_prefix| {
            reader.resolver().bindings().any(|(prefix, namespace)| {
                matches!(prefix, PrefixDeclaration::Named(value) if value == required_prefix.as_bytes())
                    && matches!(
                        namespace_kind(&ResolveResult::Bound(namespace)),
                        NamespaceKind::OfficeRelationships | NamespaceKind::Wordprocessing
                    )
            })
        }))
    }

    fn push_plain_frame(&mut self, kind: FrameKind, content_suppressed: bool) {
        self.frames.push(Frame {
            kind,
            next_child: 0,
            context_pushed: false,
            table_path: None,
            row_path: None,
            cell_path: None,
            text_box_path: None,
            content_suppressed,
        });
    }

    fn start_alternate_content(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace: NamespaceKind,
        name: &[u8],
        parent_suppressed: bool,
    ) -> Result<bool, ScanError> {
        let parent_is_alternate_content = matches!(
            self.frames.last().map(|frame| frame.kind),
            Some(FrameKind::AlternateContent { .. })
        );
        if parent_is_alternate_content {
            let eligible = namespace == NamespaceKind::MarkupCompatibility
                && match name {
                    b"Choice" => Self::alternate_choice_supported(reader, element)?,
                    b"Fallback" => true,
                    _ => false,
                };
            let parent = self.frames.last_mut().ok_or(ScanError::InvalidXml)?;
            let FrameKind::AlternateContent { branch_selected } = &mut parent.kind else {
                return Err(ScanError::InvalidXml);
            };
            let selected = !*branch_selected && eligible && !parent_suppressed;
            if selected {
                *branch_selected = true;
            }
            self.push_plain_frame(FrameKind::AlternateContentBranch, !selected);
            return Ok(true);
        }
        if namespace == NamespaceKind::MarkupCompatibility && name == b"AlternateContent" {
            self.push_plain_frame(
                FrameKind::AlternateContent {
                    branch_selected: false,
                },
                parent_suppressed,
            );
            return Ok(true);
        }
        Ok(false)
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<(), ScanError> {
        if self.frames.len() >= self.limits.depth {
            return Err(ScanError::TooDeep);
        }
        if self.frames.len() == 1 {
            if self.root_seen {
                return Err(ScanError::InvalidXml);
            }
            self.root_seen = true;
        }
        self.enter_element()?;
        let (resolved, local_name) = reader.resolver().resolve_element(element.name());
        let namespace = namespace_kind(&resolved);
        let name = local_name.as_ref();
        self.record_coverage(namespace, name)?;

        let parent_suppressed = self
            .frames
            .last()
            .ok_or(ScanError::InvalidXml)?
            .content_suppressed;
        if self.start_alternate_content(reader, element, namespace, name, parent_suppressed)? {
            return Ok(());
        }
        if parent_suppressed {
            self.push_plain_frame(FrameKind::Other, true);
            return Ok(());
        }
        if namespace != NamespaceKind::Wordprocessing {
            self.push_plain_frame(FrameKind::Other, false);
            return Ok(());
        }
        let table_path = (name == b"tbl").then(|| self.element_path.clone());
        let row_path = (name == b"tr").then(|| self.element_path.clone());
        let cell_path = (name == b"tc").then(|| self.element_path.clone());
        let text_box_path = (name == b"txbxContent").then(|| self.element_path.clone());
        let context = Self::inline_context(reader, element, name)?;
        let context_pushed = context.is_some();
        if let Some(context) = context {
            self.active_contexts.push(context);
        }

        let kind = if name == b"p" {
            let path = self.element_path.clone();
            self.begin_paragraph(&path)?
        } else if name == b"r" && self.current_block_id().is_some() {
            FrameKind::Run
        } else if matches!(name, b"t" | b"delText") {
            self.pending_text = Some(PendingText {
                block_id: self.current_block_id(),
                path: self.element_path.clone(),
                value: String::new(),
            });
            FrameKind::Text
        } else {
            if name == b"tab" && self.inside_run() {
                self.append_segment("\t", SegmentSource::Tab, self.element_path.clone())?;
            } else if matches!(name, b"br" | b"cr") && self.inside_run() {
                self.append_segment("\n", SegmentSource::Break, self.element_path.clone())?;
            }
            FrameKind::Other
        };

        self.frames.push(Frame {
            kind,
            next_child: 0,
            context_pushed,
            table_path,
            row_path,
            cell_path,
            text_box_path,
            content_suppressed: false,
        });
        Ok(())
    }

    fn append_text(&mut self, value: &str) {
        let Some(pending) = self.pending_text.as_mut() else {
            return;
        };
        pending.value.push_str(value);
    }

    fn append_segment(
        &mut self,
        value: &str,
        source: SegmentSource,
        path: Vec<usize>,
    ) -> Result<(), ScanError> {
        let Some(block_id) = self.current_block_id() else {
            return Ok(());
        };
        self.append_segment_to(block_id, value, source, path)
    }

    fn append_segment_to(
        &mut self,
        block_id: usize,
        value: &str,
        source: SegmentSource,
        path: Vec<usize>,
    ) -> Result<(), ScanError> {
        if value.is_empty() {
            return Ok(());
        }
        if self.segment_count >= self.limits.segments {
            return Err(ScanError::TooManySegments);
        }
        self.segment_count = self
            .segment_count
            .checked_add(1)
            .ok_or(ScanError::TooManySegments)?;
        let contexts = self.contexts()?;
        let builder = self
            .builders
            .get_mut(&block_id)
            .ok_or(ScanError::InvalidXml)?;
        let units = value.encode_utf16().count();
        let end_utf16 = builder
            .utf16_len
            .checked_add(units)
            .ok_or(ScanError::TextLengthOverflow)?;
        if contexts
            .iter()
            .any(|item| matches!(item, InlineContext::Hyperlink { .. }))
        {
            self.coverage.hyperlink_segments = self
                .coverage
                .hyperlink_segments
                .checked_add(1)
                .ok_or(ScanError::TooManySegments)?;
        }
        if contexts
            .iter()
            .any(|item| matches!(item, InlineContext::Revision(_)))
        {
            self.coverage.revision_segments = self
                .coverage
                .revision_segments
                .checked_add(1)
                .ok_or(ScanError::TooManySegments)?;
        }
        builder.text.push_str(value);
        builder.segments.push(TextSegment {
            start_utf16: builder.utf16_len,
            end_utf16,
            source,
            contexts,
            xml_path: path,
        });
        builder.utf16_len = end_utf16;
        Ok(())
    }

    fn end(&mut self) -> Result<(), ScanError> {
        let frame = self.frames.pop().ok_or(ScanError::InvalidXml)?;
        match frame.kind {
            FrameKind::Paragraph(block_id) => {
                let builder = self
                    .builders
                    .remove(&block_id)
                    .ok_or(ScanError::InvalidXml)?;
                if self
                    .completed
                    .insert(
                        block_id,
                        TextBlock {
                            text: builder.text,
                            location: builder.location,
                            segments: builder.segments,
                        },
                    )
                    .is_some()
                {
                    return Err(ScanError::InvalidXml);
                }
                self.flush_completed()?;
            }
            FrameKind::Text => {
                let pending = self.pending_text.take().ok_or(ScanError::InvalidXml)?;
                if !pending.value.is_empty() {
                    let block_id = pending.block_id.ok_or(ScanError::TextOutsideParagraph)?;
                    self.append_segment_to(
                        block_id,
                        &pending.value,
                        SegmentSource::Text,
                        pending.path,
                    )?;
                }
            }
            FrameKind::AlternateContent { .. }
            | FrameKind::AlternateContentBranch
            | FrameKind::Other
            | FrameKind::Run => {}
        }
        if frame.context_pushed {
            self.active_contexts.pop().ok_or(ScanError::InvalidXml)?;
        }
        self.element_path.pop().ok_or(ScanError::InvalidXml)?;
        Ok(())
    }

    fn flush_completed(&mut self) -> Result<(), ScanError> {
        while let Some(block) = self.completed.remove(&self.next_emit_id) {
            (self.sink)(block)?;
            self.next_emit_id = self
                .next_emit_id
                .checked_add(1)
                .ok_or(ScanError::TooManyBlocks)?;
        }
        Ok(())
    }

    fn finish(mut self) -> Result<PartCoverage, ScanError> {
        self.flush_completed()?;
        if !self.root_seen
            || self.frames.len() != 1
            || !self.element_path.is_empty()
            || !self.builders.is_empty()
            || !self.completed.is_empty()
            || self.pending_text.is_some()
            || !self.active_contexts.is_empty()
            || self.next_emit_id != self.next_block_id
        {
            return Err(ScanError::InvalidXml);
        }
        Ok(self.coverage)
    }
}

/// Scan one `WordprocessingML` part and deliver completed blocks in document
/// order without retaining a second whole-part block collection.
///
/// A sink may observe blocks before the scanner reaches a later malformed
/// element. Callers that require atomic behavior must stage sink output until
/// this function returns successfully.
///
/// The input boundary is deliberately UTF-8. DOCX producers conventionally
/// store `WordprocessingML` parts as UTF-8; callers with another XML encoding
/// must transcode before scanning.
///
/// # Errors
///
/// Returns [`ScanError`] when the XML is malformed, violates a configured
/// resource bound, contains text in an invalid location, or the sink rejects a
/// block.
pub fn scan_wordprocessing_part_with(
    xml: &[u8],
    limits: ScanLimits,
    sink: impl FnMut(TextBlock) -> Result<(), ScanError>,
) -> Result<PartCoverage, ScanError> {
    std::str::from_utf8(xml).map_err(|_| ScanError::UnsupportedEncoding)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner::new(limits, sink);
    loop {
        let event = reader.read_event().map_err(|_| ScanError::InvalidXml)?;
        if matches!(event, Event::Eof) {
            break;
        }
        scanner.record_event()?;
        match event {
            Event::Start(element) => scanner.start(&reader, &element)?,
            Event::Empty(element) => {
                scanner.start(&reader, &element)?;
                scanner.end()?;
            }
            Event::Text(text) => {
                let decoded = text.xml10_content().map_err(|_| ScanError::InvalidXml)?;
                scanner.append_text(&decoded);
            }
            Event::CData(text) => {
                let decoded = text.xml10_content().map_err(|_| ScanError::InvalidXml)?;
                scanner.append_text(&decoded);
            }
            Event::GeneralRef(reference) => {
                let value = if let Some(character) = reference
                    .resolve_char_ref()
                    .map_err(|_| ScanError::InvalidXml)?
                {
                    character.to_string()
                } else {
                    let name = reference.decode().map_err(|_| ScanError::InvalidXml)?;
                    resolve_predefined_entity(&name)
                        .ok_or(ScanError::InvalidXml)?
                        .to_owned()
                };
                scanner.append_text(&value);
            }
            Event::End(_) => scanner.end()?,
            Event::DocType(_) => return Err(ScanError::DocumentTypeDeclaration),
            Event::Eof => return Err(ScanError::InvalidXml),
            _ => {}
        }
    }
    scanner.finish()
}

/// Scan one `WordprocessingML` part into an owned semantic projection.
///
/// # Errors
///
/// Returns [`ScanError`] when the XML is malformed, violates a configured
/// resource bound, or contains `WordprocessingML` text outside a paragraph.
pub fn scan_wordprocessing_part(xml: &[u8], limits: ScanLimits) -> Result<PartScan, ScanError> {
    let mut blocks = Vec::new();
    let coverage = scan_wordprocessing_part_with(xml, limits, |block| {
        blocks.push(block);
        Ok(())
    })?;
    Ok(PartScan { blocks, coverage })
}

#[cfg(test)]
mod tests {
    use std::fmt::Write as _;

    use proptest::test_runner::TestCaseError;
    use proptest::{collection, prop_assert, prop_assert_eq, proptest};

    use super::{
        BlockLocation, InlineContext, RevisionKind, ScanLimits, SegmentSource,
        scan_wordprocessing_part,
    };

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    #[test]
    fn scans_locations_contexts_controls_and_utf16() -> Result<(), super::ScanError> {
        let xml = format!(
            "<x:document xmlns:x=\"{W}\" xmlns:rel=\"{R}\"><x:body><x:p><x:r><x:t>😀 </x:t></x:r><x:hyperlink rel:id=\"r1\" x:anchor=\"a\"><x:r><x:t>A</x:t></x:r></x:hyperlink><x:ins><x:r><x:t>B</x:t></x:r></x:ins><x:r><x:tab/><x:br/></x:r></x:p><x:tbl><x:tr><x:tc><x:p><x:r><x:t>C</x:t></x:r></x:p></x:tc></x:tr></x:tbl></x:body></x:document>"
        );
        let scan = scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default())?;
        assert_eq!(scan.blocks.len(), 2);
        let first = scan.blocks.first().ok_or(super::ScanError::InvalidXml)?;
        let second = scan.blocks.get(1).ok_or(super::ScanError::InvalidXml)?;
        assert_eq!(first.text, "😀 AB\t\n");
        assert_eq!(
            first
                .segments
                .first()
                .ok_or(super::ScanError::InvalidXml)?
                .end_utf16,
            3
        );
        assert_eq!(
            first
                .segments
                .get(3)
                .ok_or(super::ScanError::InvalidXml)?
                .source,
            SegmentSource::Tab
        );
        assert!(matches!(
          first
            .segments
            .get(1)
            .and_then(|segment| segment.contexts.first()),
          Some(InlineContext::Hyperlink {
            relationship_id: Some(value),
            anchor: Some(anchor),
          }) if value == "r1" && anchor == "a"
        ));
        assert!(matches!(
            first
                .segments
                .get(2)
                .and_then(|segment| segment.contexts.first()),
            Some(InlineContext::Revision(RevisionKind::Insertion))
        ));
        assert!(matches!(
            second.location,
            BlockLocation::TableCellParagraph { .. }
        ));
        Ok(())
    }

    #[test]
    fn projects_exactly_one_markup_compatibility_branch() -> Result<(), super::ScanError> {
        let unsupported_choice = format!(
            "<w:document xmlns:w=\"{W}\" xmlns:mc=\"{MC}\" xmlns:x=\"urn:unsupported\"><w:body><mc:AlternateContent><mc:Choice Requires=\"x\"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"
        );
        let fallback_scan =
            scan_wordprocessing_part(unsupported_choice.as_bytes(), ScanLimits::default())?;
        assert_eq!(
            fallback_scan
                .blocks
                .iter()
                .map(|block| block.text.as_str())
                .collect::<Vec<_>>(),
            ["fallback"]
        );
        assert_eq!(fallback_scan.coverage.unsupported_alternate_content, 1);

        let supported_choice = format!(
            "<w:document xmlns:w=\"{W}\" xmlns:mc=\"{MC}\"><w:body><mc:AlternateContent><mc:Choice Requires=\"w\"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"
        );
        let choice_scan =
            scan_wordprocessing_part(supported_choice.as_bytes(), ScanLimits::default())?;
        assert_eq!(
            choice_scan
                .blocks
                .iter()
                .map(|block| block.text.as_str())
                .collect::<Vec<_>>(),
            ["choice"]
        );
        Ok(())
    }

    #[test]
    fn emits_text_controls_only_from_run_content() -> Result<(), super::ScanError> {
        let xml = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr><w:tabs><w:tab w:val=\"left\"/></w:tabs></w:pPr><w:r><w:t>A</w:t><w:tab/><w:br/></w:r></w:p></w:body></w:document>"
        );
        let scan = scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default())?;
        let block = scan.blocks.first().ok_or(super::ScanError::InvalidXml)?;
        assert_eq!(block.text, "A\t\n");
        assert_eq!(
            block
                .segments
                .iter()
                .map(|segment| segment.source)
                .collect::<Vec<_>>(),
            [
                SegmentSource::Text,
                SegmentSource::Tab,
                SegmentSource::Break
            ]
        );
        Ok(())
    }

    #[test]
    fn decodes_xml_references_exactly_once() -> Result<(), super::ScanError> {
        let xml = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>&amp;amp; &lt; &#x1F600;</w:t></w:r></w:p></w:body></w:document>"
        );
        let scan = scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default())?;
        assert_eq!(
            scan.blocks
                .first()
                .ok_or(super::ScanError::InvalidXml)?
                .text,
            "&amp; < 😀"
        );
        Ok(())
    }

    #[test]
    fn reports_the_utf8_input_boundary_explicitly() {
        assert_eq!(
            scan_wordprocessing_part(&[0xff], ScanLimits::default()),
            Err(super::ScanError::UnsupportedEncoding)
        );
    }

    #[test]
    fn rejects_nonempty_text_bearing_elements_outside_paragraphs() -> Result<(), super::ScanError> {
        let xml = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:r><w:t>outside</w:t></w:r></w:body></w:document>"
        );
        assert_eq!(
            scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default()),
            Err(super::ScanError::TextOutsideParagraph)
        );

        let empty =
            format!("<w:document xmlns:w=\"{W}\"><w:body><w:r><w:t/></w:r></w:body></w:document>");
        let empty_scan = scan_wordprocessing_part(empty.as_bytes(), ScanLimits::default())?;
        assert!(empty_scan.blocks.is_empty());
        Ok(())
    }

    #[test]
    fn rejects_namespace_spoofing_and_accepts_strict_ooxml() -> Result<(), super::ScanError> {
        let spoof = b"<w:document xmlns:w=\"urn:not-word\"><w:body><w:p><w:r><w:t>secret</w:t></w:r></w:p></w:body></w:document>";
        let spoof_scan = scan_wordprocessing_part(spoof, ScanLimits::default())?;
        assert!(spoof_scan.blocks.is_empty());

        let strict = b"<d:document xmlns:d=\"http://purl.oclc.org/ooxml/wordprocessingml/main\"><d:body><d:p><d:r><d:t>strict</d:t></d:r></d:p></d:body></d:document>";
        let strict_scan = scan_wordprocessing_part(strict, ScanLimits::default())?;
        assert_eq!(
            strict_scan
                .blocks
                .first()
                .ok_or(super::ScanError::InvalidXml)?
                .text,
            "strict"
        );
        Ok(())
    }

    #[test]
    fn nested_namespace_rebinding_cannot_spoof_wordprocessing_elements()
    -> Result<(), super::ScanError> {
        let xml = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><x xmlns:w=\"urn:not-word\"><w:p><w:r><w:t>hidden</w:t></w:r></w:p></x><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"
        );
        let scan = scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default())?;
        let texts = scan
            .blocks
            .iter()
            .map(|block| block.text.as_str())
            .collect::<Vec<_>>();
        assert_eq!(texts, ["before", "after"]);
        Ok(())
    }

    #[test]
    fn rejects_document_type_declarations() {
        let doctype = format!("<!DOCTYPE w:document><w:document xmlns:w=\"{W}\"/>");
        assert_eq!(
            scan_wordprocessing_part(doctype.as_bytes(), ScanLimits::default()),
            Err(super::ScanError::DocumentTypeDeclaration)
        );
    }

    #[test]
    fn rejects_multiple_roots_and_resource_limit_overruns() {
        let multiple_roots = format!("<w:p xmlns:w=\"{W}\"/><w:p xmlns:w=\"{W}\"/>");
        assert_eq!(
            scan_wordprocessing_part(multiple_roots.as_bytes(), ScanLimits::default()),
            Err(super::ScanError::InvalidXml)
        );
        let two_blocks =
            format!("<w:document xmlns:w=\"{W}\"><w:body><w:p/><w:p/></w:body></w:document>");
        assert_eq!(
            scan_wordprocessing_part(
                two_blocks.as_bytes(),
                ScanLimits {
                    blocks: 1,
                    ..ScanLimits::default()
                }
            ),
            Err(super::ScanError::TooManyBlocks)
        );

        let empty_document = format!("<w:document xmlns:w=\"{W}\"/>");
        let one_event = scan_wordprocessing_part(
            empty_document.as_bytes(),
            ScanLimits {
                events: 1,
                ..ScanLimits::default()
            },
        );
        assert!(matches!(
            one_event,
            Ok(super::PartScan {
                coverage: super::PartCoverage {
                    work: super::ScanWork { events: 1, .. },
                    ..
                },
                ..
            })
        ));

        let one_block = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"
        );
        assert_eq!(
            scan_wordprocessing_part(
                one_block.as_bytes(),
                ScanLimits {
                    events: 1,
                    ..ScanLimits::default()
                }
            ),
            Err(super::ScanError::TooManyEvents)
        );

        let contextual = format!(
            "<w:document xmlns:w=\"{W}\"><w:body><w:p><w:ins><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body></w:document>"
        );
        assert_eq!(
            scan_wordprocessing_part(
                contextual.as_bytes(),
                ScanLimits {
                    inline_context_copies: 0,
                    ..ScanLimits::default()
                }
            ),
            Err(super::ScanError::TooManyInlineContextCopies)
        );
    }

    fn sibling_run_fixture(run_count: usize) -> String {
        let mut runs = String::new();
        for _ in 0..run_count {
            runs.push_str("<w:r><w:t>x</w:t></w:r>");
        }
        format!("<w:document xmlns:w=\"{W}\"><w:body><w:p>{runs}</w:p></w:body></w:document>")
    }

    fn sibling_paragraph_fixture(paragraph_count: usize) -> String {
        let mut paragraphs = String::new();
        for _ in 0..paragraph_count {
            paragraphs.push_str("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
        }
        format!("<w:document xmlns:w=\"{W}\"><w:body>{paragraphs}</w:body></w:document>")
    }

    #[test]
    fn scan_work_is_additive_across_independent_growth_axes() -> Result<(), super::ScanError> {
        let empty_runs =
            scan_wordprocessing_part(sibling_run_fixture(0).as_bytes(), ScanLimits::default())?;
        let small_runs =
            scan_wordprocessing_part(sibling_run_fixture(512).as_bytes(), ScanLimits::default())?;
        let large_runs =
            scan_wordprocessing_part(sibling_run_fixture(1_024).as_bytes(), ScanLimits::default())?;
        assert_eq!(
            large_runs.coverage.work.events - empty_runs.coverage.work.events,
            2 * (small_runs.coverage.work.events - empty_runs.coverage.work.events),
        );
        assert_eq!(small_runs.coverage.work.inline_context_copies, 0);
        assert_eq!(large_runs.coverage.work.inline_context_copies, 0);

        let empty_paragraphs = scan_wordprocessing_part(
            sibling_paragraph_fixture(0).as_bytes(),
            ScanLimits::default(),
        )?;
        let small_paragraphs = scan_wordprocessing_part(
            sibling_paragraph_fixture(256).as_bytes(),
            ScanLimits::default(),
        )?;
        let large_paragraphs = scan_wordprocessing_part(
            sibling_paragraph_fixture(512).as_bytes(),
            ScanLimits::default(),
        )?;
        assert_eq!(
            large_paragraphs.coverage.work.events - empty_paragraphs.coverage.work.events,
            2 * (small_paragraphs.coverage.work.events - empty_paragraphs.coverage.work.events),
        );
        Ok(())
    }

    proptest! {
      #[test]
      fn arbitrary_prefixes_and_run_splits_preserve_text_and_utf16_offsets(
        prefix in "[a-z]{1,8}",
        fragments in collection::vec("[A-Za-z0-9 ]{0,16}", 1..32),
      ) {
        let mut runs = String::new();
        for fragment in &fragments {
          write!(
            &mut runs,
            "<{prefix}:r><{prefix}:t>{fragment}</{prefix}:t></{prefix}:r>"
          )
          .map_err(|error| TestCaseError::fail(error.to_string()))?;
        }
        let xml = format!(
          "<{prefix}:document xmlns:{prefix}=\"{W}\"><{prefix}:body><{prefix}:p>{runs}</{prefix}:p></{prefix}:body></{prefix}:document>"
        );
        let scan = scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default())
          .map_err(|error| TestCaseError::fail(error.to_string()))?;
        let block = scan
          .blocks
          .first()
          .ok_or_else(|| TestCaseError::fail("missing projected paragraph"))?;
        prop_assert_eq!(&block.text, &fragments.concat());
        let mut expected_start = 0_usize;
        for segment in &block.segments {
          prop_assert_eq!(segment.start_utf16, expected_start);
          prop_assert!(segment.end_utf16 >= segment.start_utf16);
          expected_start = segment.end_utf16;
        }
        prop_assert_eq!(expected_start, block.text.encode_utf16().count());
      }

      #[test]
      fn nonempty_wordprocessing_text_requires_a_paragraph(
        prefix in "[a-z]{1,8}",
        namespace in proptest::sample::select(&[W, W_STRICT][..]),
        text_element in proptest::sample::select(&["t", "delText"][..]),
        value in "[A-Za-z0-9 ]{1,32}",
        run_wrapped in proptest::bool::ANY,
      ) {
        let text = format!(
          "<{prefix}:{text_element}>{value}</{prefix}:{text_element}>"
        );
        let content = if run_wrapped {
          format!("<{prefix}:r>{text}</{prefix}:r>")
        } else {
          text
        };
        let xml = format!(
          "<{prefix}:document xmlns:{prefix}=\"{namespace}\"><{prefix}:body>{content}</{prefix}:body></{prefix}:document>"
        );
        prop_assert_eq!(
          scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default()),
          Err(super::ScanError::TextOutsideParagraph)
        );
      }

      #[test]
      fn document_type_declarations_have_a_stable_typed_error(
        document_type in "[A-Za-z][A-Za-z0-9]{0,15}",
      ) {
        let xml = format!(
          "<!DOCTYPE {document_type}><w:document xmlns:w=\"{W}\"/>"
        );
        prop_assert_eq!(
          scan_wordprocessing_part(xml.as_bytes(), ScanLimits::default()),
          Err(super::ScanError::DocumentTypeDeclaration)
        );
      }
    }
}
