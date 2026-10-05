//! Borrowed typed views over the existing model JSON. Untouched fields stay verbatim.
//!
//! This spike deliberately does not claim to port the complete OOXML model. The
//! views type the records the implemented operations inspect; the backing value
//! retains extension fields, opaque markup, and property spelling.

use serde_json::{Map, Value};
use std::collections::BTreeSet;
use std::fmt;

pub type FieldPath = Vec<String>;

/// JSON alone cannot express an owned JavaScript `undefined` property.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FieldPresence<'a> {
    Absent,
    Null,
    Undefined,
    Value(&'a Value),
}

/// Read an owned field without treating an omitted key as a null value.
pub fn field_presence<'a>(
    value: &'a Value,
    undefined_fields: &BTreeSet<FieldPath>,
    path: &[String],
) -> FieldPresence<'a> {
    if undefined_fields.contains(path) {
        return FieldPresence::Undefined;
    }
    match value_at_path(value, path) {
        None => FieldPresence::Absent,
        Some(Value::Null) => FieldPresence::Null,
        Some(value) => FieldPresence::Value(value),
    }
}

pub fn array_index(field: &str) -> Option<usize> {
    if field.is_empty() || (field.len() > 1 && field.starts_with('0')) {
        return None;
    }
    if !field.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    field.parse().ok()
}

pub fn value_at_path<'a>(mut value: &'a Value, path: &[String]) -> Option<&'a Value> {
    for field in path {
        value = match value {
            Value::Object(object) => object.get(field)?,
            Value::Array(array) => array.get(array_index(field)?)?,
            _ => return None,
        };
    }
    Some(value)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModelError {
    pub message: String,
}

impl ModelError {
    fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl fmt::Display for ModelError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for ModelError {}

#[derive(Debug, Clone, PartialEq)]
pub struct DocumentModel {
    pub value: Value,
    /// Document presence is a sidecar, never encoded as a null placeholder.
    pub undefined_fields: BTreeSet<FieldPath>,
}

impl DocumentModel {
    pub fn from_json(value: Value) -> Result<Self, ModelError> {
        Self::from_parts(value, BTreeSet::new())
    }

    pub fn from_parts(
        value: Value,
        undefined_fields: BTreeSet<FieldPath>,
    ) -> Result<Self, ModelError> {
        let model = Self {
            value,
            undefined_fields,
        };
        model.body()?;
        Ok(model)
    }

    pub fn body(&self) -> Result<DocumentBodyView<'_>, ModelError> {
        let body = self
            .value
            .get("package")
            .and_then(|package| package.get("document"))
            .ok_or_else(|| ModelError::new("A document must contain package.document."))?;
        DocumentBodyView::from_json(body)
    }

    pub fn presence(&self, path: &[String]) -> FieldPresence<'_> {
        field_presence(&self.value, &self.undefined_fields, path)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OpStory {
    Main,
    Header { r_id: String },
    Footer { r_id: String },
    Footnote { id: i64 },
    Endnote { id: i64 },
}

impl OpStory {
    pub fn from_json(value: &Value) -> Result<Self, ModelError> {
        if value.as_str() == Some("main") {
            return Ok(Self::Main);
        }
        let object = value
            .as_object()
            .ok_or_else(|| ModelError::new("A story must be main or an addressed story object."))?;
        match object.get("kind").and_then(Value::as_str) {
            Some("header") => Ok(Self::Header {
                r_id: required_string(object, "rId")?.to_owned(),
            }),
            Some("footer") => Ok(Self::Footer {
                r_id: required_string(object, "rId")?.to_owned(),
            }),
            Some("footnote") => Ok(Self::Footnote {
                id: required_note_id(object)?,
            }),
            Some("endnote") => Ok(Self::Endnote {
                id: required_note_id(object)?,
            }),
            _ => Err(ModelError::new("The story kind is unsupported.")),
        }
    }

    pub fn to_json(&self) -> Value {
        match self {
            Self::Main => Value::String("main".to_owned()),
            Self::Header { r_id } => serde_json::json!({"kind":"header", "rId":r_id}),
            Self::Footer { r_id } => serde_json::json!({"kind":"footer", "rId":r_id}),
            Self::Footnote { id } => serde_json::json!({"kind":"footnote", "id":id}),
            Self::Endnote { id } => serde_json::json!({"kind":"endnote", "id":id}),
        }
    }
}

fn required_note_id(object: &Map<String, Value>) -> Result<i64, ModelError> {
    object
        .get("id")
        .and_then(Value::as_i64)
        .ok_or_else(|| ModelError::new("A note story id must be an integer."))
}

fn required_string<'a>(object: &'a Map<String, Value>, field: &str) -> Result<&'a str, ModelError> {
    object
        .get(field)
        .and_then(Value::as_str)
        .ok_or_else(|| ModelError::new(format!("The {field} field must be a string.")))
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextPosition {
    pub story: OpStory,
    pub block_id: String,
    pub offset: usize,
    pub zero_width_before: Option<usize>,
}

impl TextPosition {
    pub fn from_json(value: &Value) -> Result<Self, ModelError> {
        let object = value
            .as_object()
            .ok_or_else(|| ModelError::new("A text position must be an object."))?;
        let story = OpStory::from_json(
            object
                .get("story")
                .ok_or_else(|| ModelError::new("A text position must name its story."))?,
        )?;
        let offset = required_offset(object.get("offset"), "offset")?;
        let zero_width_before = object
            .get("zeroWidthBefore")
            .map(|value| required_offset(Some(value), "zeroWidthBefore"))
            .transpose()?;
        Ok(Self {
            story,
            block_id: required_string(object, "blockId")?.to_owned(),
            offset,
            zero_width_before,
        })
    }
}

fn required_offset(value: Option<&Value>, field: &str) -> Result<usize, ModelError> {
    value
        .and_then(Value::as_u64)
        .and_then(|number| usize::try_from(number).ok())
        .ok_or_else(|| ModelError::new(format!("The {field} field must be a nonnegative integer.")))
}

#[derive(Debug, Clone, Copy)]
pub struct DocumentBodyView<'a> {
    pub value: &'a Value,
    pub content: &'a [Value],
}

impl<'a> DocumentBodyView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let content = value
            .as_object()
            .and_then(|object| object.get("content"))
            .and_then(Value::as_array)
            .ok_or_else(|| ModelError::new("A document body must contain a content array."))?;
        Ok(Self { value, content })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ParagraphView<'a> {
    pub value: &'a Value,
    pub para_id: Option<&'a str>,
    pub content: &'a [Value],
}

impl<'a> ParagraphView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let object = typed_object(value, "paragraph")?;
        let content = required_content(object)?;
        let para_id = object
            .get("paraId")
            .map(|id| {
                id.as_str()
                    .ok_or_else(|| ModelError::new("A paragraph id must be a string."))
            })
            .transpose()?;
        Ok(Self {
            value,
            para_id,
            content,
        })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct RunView<'a> {
    pub value: &'a Value,
    pub content: &'a [Value],
}

impl<'a> RunView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        Ok(Self {
            value,
            content: required_content(typed_object(value, "run")?)?,
        })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct TextView<'a> {
    pub value: &'a Value,
    pub text: &'a str,
}

impl<'a> TextView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        Ok(Self {
            value,
            text: required_string(typed_object(value, "text")?, "text")?,
        })
    }
}

fn typed_object<'a>(value: &'a Value, kind: &str) -> Result<&'a Map<String, Value>, ModelError> {
    let object = value
        .as_object()
        .ok_or_else(|| ModelError::new(format!("A {kind} record must be an object.")))?;
    if object.get("type").and_then(Value::as_str) != Some(kind) {
        return Err(ModelError::new(format!("Expected a {kind} record.")));
    }
    Ok(object)
}

fn required_content(object: &Map<String, Value>) -> Result<&[Value], ModelError> {
    object
        .get("content")
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .ok_or_else(|| ModelError::new("A content record must contain a content array."))
}

/// Typed topology view; opaque row and cell metadata remain on their backing records.
#[derive(Debug, Clone, Copy)]
pub struct TableView<'a> {
    pub value: &'a Value,
    pub rows: &'a [Value],
}

impl<'a> TableView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let fields = typed_object(value, "table")?;
        let rows = fields
            .get("rows")
            .and_then(Value::as_array)
            .ok_or_else(|| ModelError::new("A table needs a rows array."))?;
        Ok(Self { value, rows })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct TableRowView<'a> {
    pub value: &'a Value,
    pub cells: &'a [Value],
}

impl<'a> TableRowView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let fields = typed_object(value, "tableRow")?;
        let cells = fields
            .get("cells")
            .and_then(Value::as_array)
            .ok_or_else(|| ModelError::new("A table row needs a cells array."))?;
        Ok(Self { value, cells })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct TableCellView<'a> {
    pub value: &'a Value,
    pub content: &'a [Value],
}

impl<'a> TableCellView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let fields = typed_object(value, "tableCell")?;
        Ok(Self {
            value,
            content: required_content(fields)?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrackedKind {
    Insertion,
    Deletion,
}

#[derive(Debug, Clone, Copy)]
pub struct TrackedView<'a> {
    pub value: &'a Value,
    pub kind: TrackedKind,
    pub info: &'a Value,
    pub content: &'a [Value],
}

impl<'a> TrackedView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let kind = match value.get("type").and_then(Value::as_str) {
            Some("insertion") => TrackedKind::Insertion,
            Some("deletion") => TrackedKind::Deletion,
            _ => return Err(ModelError::new("Expected insertion or deletion.")),
        };
        let fields = value
            .as_object()
            .ok_or_else(|| ModelError::new("A tracked record must be an object."))?;
        let info = fields
            .get("info")
            .filter(|value| value.is_object())
            .ok_or_else(|| ModelError::new("A tracked record needs info."))?;
        Ok(Self {
            value,
            kind,
            info,
            content: required_content(fields)?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CommentAnchorKind {
    RangeStart,
    RangeEnd,
    Reference,
}

#[derive(Debug, Clone, Copy)]
pub struct CommentAnchorView<'a> {
    pub value: &'a Value,
    pub kind: CommentAnchorKind,
    pub id: i64,
}

impl<'a> CommentAnchorView<'a> {
    pub fn from_json(value: &'a Value) -> Result<Self, ModelError> {
        let kind = match value.get("type").and_then(Value::as_str) {
            Some("commentRangeStart") => CommentAnchorKind::RangeStart,
            Some("commentRangeEnd") => CommentAnchorKind::RangeEnd,
            Some("commentReference") => CommentAnchorKind::Reference,
            _ => return Err(ModelError::new("Expected a comment anchor.")),
        };
        let id = value
            .get("id")
            .and_then(Value::as_i64)
            .ok_or_else(|| ModelError::new("A comment anchor needs an integer id."))?;
        Ok(Self { value, kind, id })
    }
}
