use std::collections::{HashMap, HashSet};

use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const WORDPROCESSING_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORDPROCESSING_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const WORDPROCESSING_2010: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const WORDPROCESSING_2012: &[u8] = b"http://schemas.microsoft.com/office/word/2012/wordml";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReviewFactLimits {
    pub maximum_comments_xml_bytes: usize,
    pub maximum_comments_extended_xml_bytes: usize,
    pub maximum_facts_per_family: usize,
}

impl Default for ReviewFactLimits {
    fn default() -> Self {
        Self {
            maximum_comments_xml_bytes: 16 * 1024 * 1024,
            maximum_comments_extended_xml_bytes: 16 * 1024 * 1024,
            maximum_facts_per_family: 1_000_000,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReviewFactUnknownReason {
    InvalidDocument,
    InvalidComments,
    InvalidCommentsExtended,
    ResourceLimit,
    UnsupportedLocation,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ReviewFactSet<T> {
    Known(Vec<T>),
    Unknown(ReviewFactUnknownReason),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ReviewDetail<T> {
    Known(T),
    Unknown(ReviewFactUnknownReason),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReviewPoint {
    pub paragraph_ordinal: usize,
    pub utf8: u32,
    pub utf16: u32,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReviewSpan {
    pub start: ReviewPoint,
    pub end: ReviewPoint,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct RevisionContent {
    pub span: ReviewSpan,
    pub text: String,
    pub formatting_only: bool,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CommentContent {
    pub anchor: ReviewSpan,
    pub comment_text: String,
    pub referenced_text: String,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RevisionFactKind {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    CellInsertion,
    CellDeletion,
    CellMerge,
    ParagraphPropertiesChange,
    RunPropertiesChange,
    SectionPropertiesChange,
    TablePropertiesChange,
    TableRowPropertiesChange,
    TableCellPropertiesChange,
    TableGridChange,
    CustomXmlDeletionRangeStart,
    CustomXmlDeletionRangeEnd,
    CustomXmlInsertionRangeStart,
    CustomXmlInsertionRangeEnd,
    CustomXmlMoveFromRangeStart,
    CustomXmlMoveFromRangeEnd,
    CustomXmlMoveToRangeStart,
    CustomXmlMoveToRangeEnd,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct AttributedRevision {
    pub kind: RevisionFactKind,
    pub author: String,
    pub date: Option<String>,
    pub revision_id: Option<String>,
    pub content: ReviewDetail<RevisionContent>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct AttributedComment {
    pub comment_id: String,
    pub author: String,
    pub initials: Option<String>,
    pub date: Option<String>,
    pub parent_comment_id: Option<String>,
    pub resolved: bool,
    pub content: ReviewDetail<CommentContent>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DocumentReviewFacts {
    pub revisions: ReviewFactSet<AttributedRevision>,
    pub comments: ReviewFactSet<AttributedComment>,
}

#[derive(Clone)]
struct CommentRow {
    comment_id: String,
    author: String,
    initials: Option<String>,
    date: Option<String>,
    paragraph_id: Option<String>,
    paragraph_id_from_wrapper: bool,
}

#[derive(Clone)]
struct CommentExtension {
    parent_paragraph_id: Option<String>,
    resolved: bool,
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum NamespaceKind {
    Other,
    Wordprocessing,
    Wordprocessing2010,
    Wordprocessing2012,
}

pub(super) fn project_review_facts(
    revisions: ReviewFactSet<AttributedRevision>,
    comments_xml: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
    comments_extended_xml: Result<Option<Vec<u8>>, ReviewFactUnknownReason>,
    limits: ReviewFactLimits,
) -> DocumentReviewFacts {
    let comments = match comments_xml {
        Ok(None) => ReviewFactSet::Known(Vec::new()),
        Ok(Some(comments_xml)) => match comments_extended_xml {
            Ok(comments_extended_xml) => parse_comments(
                &comments_xml,
                comments_extended_xml.as_deref(),
                limits.maximum_facts_per_family,
            )
            .map_or_else(ReviewFactSet::Unknown, ReviewFactSet::Known),
            Err(reason) => ReviewFactSet::Unknown(reason),
        },
        Err(reason) => ReviewFactSet::Unknown(reason),
    };
    DocumentReviewFacts {
        revisions,
        comments,
    }
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    let ResolveResult::Bound(Namespace(value)) = namespace else {
        return NamespaceKind::Other;
    };
    match *value {
        WORDPROCESSING_TRANSITIONAL | WORDPROCESSING_STRICT => NamespaceKind::Wordprocessing,
        WORDPROCESSING_2010 => NamespaceKind::Wordprocessing2010,
        WORDPROCESSING_2012 => NamespaceKind::Wordprocessing2012,
        _ => NamespaceKind::Other,
    }
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    local_name: &[u8],
) -> Result<Option<String>, ReviewFactUnknownReason> {
    for item in element.attributes() {
        let item = item.map_err(|_| ReviewFactUnknownReason::InvalidDocument)?;
        let (resolved, name) = reader.resolver().resolve_attribute(item.key);
        if namespace_kind(&resolved) == namespace && name.as_ref() == local_name {
            return item
                .decoded_and_normalized_value(XmlVersion::default(), reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|_| ReviewFactUnknownReason::InvalidDocument);
        }
    }
    Ok(None)
}

fn comment_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    local_name: &[u8],
) -> Result<Option<String>, ReviewFactUnknownReason> {
    attribute(reader, element, namespace, local_name)
        .map_err(|_| ReviewFactUnknownReason::InvalidComments)
}

fn comment_extension_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>, ReviewFactUnknownReason> {
    attribute(
        reader,
        element,
        NamespaceKind::Wordprocessing2012,
        local_name,
    )
    .map_err(|_| ReviewFactUnknownReason::InvalidCommentsExtended)
}

fn comment_paragraph_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<String>, ReviewFactUnknownReason> {
    for namespace in [
        NamespaceKind::Wordprocessing2010,
        NamespaceKind::Wordprocessing2012,
        NamespaceKind::Wordprocessing,
    ] {
        if let Some(value) = comment_attribute(reader, element, namespace, b"paraId")? {
            return Ok(Some(value.to_ascii_uppercase()));
        }
    }
    Ok(None)
}

fn parse_comments(
    comments_xml: &[u8],
    comments_extended_xml: Option<&[u8]>,
    maximum_facts: usize,
) -> Result<Vec<AttributedComment>, ReviewFactUnknownReason> {
    let comments = parse_comment_rows(comments_xml, maximum_facts)?;
    let Some(comments_extended_xml) = comments_extended_xml else {
        return Ok(comments
            .into_iter()
            .map(|comment| attributed_comment(comment, None, false))
            .collect());
    };
    let mut extensions = parse_comment_extensions(comments_extended_xml, maximum_facts)?;
    let comment_ids_by_paragraph = comments
        .iter()
        .filter_map(|comment| {
            comment
                .paragraph_id
                .as_ref()
                .map(|paragraph_id| (paragraph_id.clone(), comment.comment_id.clone()))
        })
        .collect::<HashMap<_, _>>();
    let comments_with_paragraph_ids = comments
        .iter()
        .filter(|comment| comment.paragraph_id.is_some())
        .count();
    if comment_ids_by_paragraph.len() != comments_with_paragraph_ids {
        return Err(ReviewFactUnknownReason::InvalidCommentsExtended);
    }
    let mut output = Vec::with_capacity(comments.len());
    for comment in comments {
        let extension = comment
            .paragraph_id
            .as_ref()
            .and_then(|paragraph_id| extensions.remove(paragraph_id));
        let (parent_comment_id, resolved) = extension.map_or(Ok((None, false)), |extension| {
            let parent = extension
                .parent_paragraph_id
                .as_ref()
                .map(|parent| {
                    comment_ids_by_paragraph
                        .get(parent)
                        .cloned()
                        .ok_or(ReviewFactUnknownReason::InvalidCommentsExtended)
                })
                .transpose()?;
            Ok((parent, extension.resolved))
        })?;
        output.push(attributed_comment(comment, parent_comment_id, resolved));
    }
    if !extensions.is_empty() || contains_parent_cycle(&output) {
        return Err(ReviewFactUnknownReason::InvalidCommentsExtended);
    }
    Ok(output)
}

fn attributed_comment(
    comment: CommentRow,
    parent_comment_id: Option<String>,
    resolved: bool,
) -> AttributedComment {
    AttributedComment {
        comment_id: comment.comment_id,
        author: comment.author,
        initials: comment.initials,
        date: comment.date,
        parent_comment_id,
        resolved,
        content: ReviewDetail::Unknown(ReviewFactUnknownReason::UnsupportedLocation),
    }
}

#[allow(clippy::too_many_lines)] // One event loop keeps comment shape and depth validation atomic.
fn parse_comment_rows(
    xml: &[u8],
    maximum_facts: usize,
) -> Result<Vec<CommentRow>, ReviewFactUnknownReason> {
    let invalid = ReviewFactUnknownReason::InvalidComments;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().expand_empty_elements = true;
    reader.config_mut().check_end_names = true;
    let mut output = Vec::new();
    let mut current: Option<(usize, CommentRow)> = None;
    let mut depth = 0_usize;
    let mut root_seen = false;
    let mut comment_ids = HashSet::new();
    let mut paragraph_ids = HashSet::new();
    loop {
        match reader.read_resolved_event() {
            Ok((_, Event::Eof)) => {
                if depth != 0 || current.is_some() || !root_seen {
                    return Err(invalid);
                }
                return Ok(output);
            }
            Ok((namespace, Event::Start(element))) => {
                let kind = namespace_kind(&namespace);
                if depth == 0 {
                    if root_seen
                        || kind != NamespaceKind::Wordprocessing
                        || element.local_name().as_ref() != b"comments"
                    {
                        return Err(invalid);
                    }
                    root_seen = true;
                } else if depth == 1
                    && kind == NamespaceKind::Wordprocessing
                    && element.local_name().as_ref() == b"comment"
                {
                    if current.is_some() || output.len() >= maximum_facts {
                        return Err(if output.len() >= maximum_facts {
                            ReviewFactUnknownReason::ResourceLimit
                        } else {
                            invalid
                        });
                    }
                    let comment_id =
                        comment_attribute(&reader, &element, NamespaceKind::Wordprocessing, b"id")?
                            .ok_or(invalid)?;
                    if !comment_ids.insert(comment_id.clone()) {
                        return Err(invalid);
                    }
                    let wrapper_paragraph_id = comment_paragraph_id(&reader, &element)?;
                    let paragraph_id_from_wrapper = wrapper_paragraph_id.is_some();
                    current = Some((
                        depth,
                        CommentRow {
                            comment_id,
                            author: comment_attribute(
                                &reader,
                                &element,
                                NamespaceKind::Wordprocessing,
                                b"author",
                            )?
                            .unwrap_or_default(),
                            initials: comment_attribute(
                                &reader,
                                &element,
                                NamespaceKind::Wordprocessing,
                                b"initials",
                            )?,
                            date: comment_attribute(
                                &reader,
                                &element,
                                NamespaceKind::Wordprocessing,
                                b"date",
                            )?,
                            paragraph_id: wrapper_paragraph_id,
                            paragraph_id_from_wrapper,
                        },
                    ));
                } else if kind == NamespaceKind::Wordprocessing
                    && element.local_name().as_ref() == b"p"
                    && let Some((comment_depth, comment)) = current.as_mut()
                    && depth == comment_depth.saturating_add(1)
                {
                    // Exporters may put the join key on the wrapper. Otherwise
                    // Word keys commentsExtended by the last direct-child paragraph.
                    if !comment.paragraph_id_from_wrapper {
                        comment.paragraph_id = comment_paragraph_id(&reader, &element)?;
                    }
                }
                depth = depth.checked_add(1).ok_or(invalid)?;
            }
            Ok((_, Event::DocType(_))) | Err(_) => return Err(invalid),
            Ok((namespace, Event::End(element))) => {
                depth = depth.checked_sub(1).ok_or(invalid)?;
                if namespace_kind(&namespace) == NamespaceKind::Wordprocessing
                    && element.local_name().as_ref() == b"comment"
                {
                    let (comment_depth, comment) = current.take().ok_or(invalid)?;
                    if depth != comment_depth {
                        return Err(invalid);
                    }
                    if let Some(paragraph_id) = &comment.paragraph_id
                        && !paragraph_ids.insert(paragraph_id.clone())
                    {
                        return Err(invalid);
                    }
                    output.push(comment);
                }
            }
            Ok(_) => {}
        }
    }
}

fn parse_comment_extensions(
    xml: &[u8],
    maximum_facts: usize,
) -> Result<HashMap<String, CommentExtension>, ReviewFactUnknownReason> {
    let invalid = ReviewFactUnknownReason::InvalidCommentsExtended;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().expand_empty_elements = true;
    reader.config_mut().check_end_names = true;
    let mut output = HashMap::new();
    let mut depth = 0_usize;
    let mut root_seen = false;
    loop {
        match reader.read_resolved_event() {
            Ok((_, Event::Eof)) => {
                if depth != 0 || !root_seen {
                    return Err(invalid);
                }
                return Ok(output);
            }
            Ok((namespace, Event::Start(element))) => {
                let kind = namespace_kind(&namespace);
                if depth == 0 {
                    if root_seen
                        || kind != NamespaceKind::Wordprocessing2012
                        || element.local_name().as_ref() != b"commentsEx"
                    {
                        return Err(invalid);
                    }
                    root_seen = true;
                } else if depth == 1
                    && kind == NamespaceKind::Wordprocessing2012
                    && element.local_name().as_ref() == b"commentEx"
                {
                    if output.len() >= maximum_facts {
                        return Err(ReviewFactUnknownReason::ResourceLimit);
                    }
                    let paragraph_id = comment_extension_attribute(&reader, &element, b"paraId")?
                        .ok_or(invalid)?
                        .to_ascii_uppercase();
                    let parent_paragraph_id =
                        comment_extension_attribute(&reader, &element, b"paraIdParent")?
                            .map(|value| value.to_ascii_uppercase());
                    let resolved = comment_extension_attribute(&reader, &element, b"done")?
                        .map(|value| parse_on_off(&value))
                        .transpose()?
                        .unwrap_or(false);
                    if output
                        .insert(
                            paragraph_id,
                            CommentExtension {
                                parent_paragraph_id,
                                resolved,
                            },
                        )
                        .is_some()
                    {
                        return Err(invalid);
                    }
                }
                depth = depth.checked_add(1).ok_or(invalid)?;
            }
            Ok((_, Event::End(_))) => depth = depth.checked_sub(1).ok_or(invalid)?,
            Ok((_, Event::DocType(_))) | Err(_) => return Err(invalid),
            Ok(_) => {}
        }
    }
}

fn parse_on_off(value: &str) -> Result<bool, ReviewFactUnknownReason> {
    match value.to_ascii_lowercase().as_str() {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(ReviewFactUnknownReason::InvalidCommentsExtended),
    }
}

fn contains_parent_cycle(comments: &[AttributedComment]) -> bool {
    let parents = comments
        .iter()
        .map(|comment| {
            (
                comment.comment_id.as_str(),
                comment.parent_comment_id.as_deref(),
            )
        })
        .collect::<HashMap<_, _>>();
    let mut complete = HashSet::new();
    for comment_id in parents.keys() {
        let mut path = Vec::new();
        let mut visiting = HashSet::new();
        let mut current = Some(*comment_id);
        while let Some(id) = current {
            if complete.contains(id) {
                break;
            }
            if !visiting.insert(id) {
                return true;
            }
            path.push(id);
            current = parents.get(id).copied().flatten();
        }
        complete.extend(path);
    }
    false
}
