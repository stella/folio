//! Schema-9 operation JSON and the existing `undefinedFields` presence sidecar.

use crate::model::{FieldPath, FieldPresence, array_index, field_presence, value_at_path};
use serde_json::Value;
use std::collections::BTreeSet;
use std::fmt;

pub const DOCUMENT_OP_SCHEMA_VERSION: u64 = 9;
pub const PRESENCE_FIELD: &str = "undefinedFields";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WireError {
    pub reason: &'static str,
    pub message: String,
    pub op_type: Option<String>,
}

impl WireError {
    fn structure(value: &Value, message: impl Into<String>) -> Self {
        Self {
            reason: "structureMismatch",
            message: message.into(),
            op_type: value.get("type").and_then(Value::as_str).map(str::to_owned),
        }
    }
}

impl fmt::Display for WireError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for WireError {}

#[derive(Debug, Clone, PartialEq)]
pub struct CapturedOperation {
    /// Original JSON spelling, including presence metadata and unknown fields.
    pub value: Value,
    pub undefined_fields: BTreeSet<FieldPath>,
}

impl CapturedOperation {
    pub fn presence(&self, path: &[String]) -> FieldPresence<'_> {
        field_presence(&self.value, &self.undefined_fields, path)
    }

    /// Enumerate owned patch fields, including keys JSON omitted as undefined.
    pub fn patch_fields(
        &self,
        path: &[String],
    ) -> Result<Vec<(String, FieldPresence<'_>)>, WireError> {
        let parent = value_at_path(&self.value, path)
            .and_then(Value::as_object)
            .ok_or_else(|| {
                WireError::structure(&self.value, "A formatting patch must be an object.")
            })?;
        let mut keys = parent.keys().cloned().collect::<BTreeSet<_>>();
        for undefined in &self.undefined_fields {
            if undefined.len() == path.len() + 1 && undefined.starts_with(path) {
                keys.insert(undefined[path.len()].clone());
            }
        }
        Ok(keys
            .into_iter()
            .map(|key| {
                let mut child = path.to_vec();
                child.push(key.clone());
                (key, self.presence(&child))
            })
            .collect())
    }
}

pub fn decode_envelope(
    value: &Value,
    accepted_schema: u64,
) -> Result<CapturedOperation, WireError> {
    let object = value.as_object().ok_or_else(|| {
        WireError::structure(value, "A document operation envelope must be an object.")
    })?;
    if object.get("schema").and_then(Value::as_u64) != Some(accepted_schema) {
        return Err(WireError {
            reason: "unsupportedSchema",
            message: format!("The operation envelope must use schema {accepted_schema}."),
            op_type: object
                .get("op")
                .and_then(|op| op.get("type"))
                .and_then(Value::as_str)
                .map(str::to_owned),
        });
    }
    let op = object.get("op").ok_or_else(|| {
        WireError::structure(
            value,
            "A document operation envelope must contain an operation.",
        )
    })?;
    validate_operation(op)
}

pub fn encode_envelope(op: &CapturedOperation, schema: u64) -> Value {
    serde_json::json!({ "schema": schema, "op": op.value })
}

/// Validate metadata under the same ownership and conflict rules as TS restoreDocumentOp.
/// `undefined` remains an internal state; it is never replaced by JSON null.
pub fn validate_operation(value: &Value) -> Result<CapturedOperation, WireError> {
    let object = value
        .as_object()
        .ok_or_else(|| WireError::structure(value, "A document operation must be an object."))?;
    let mut undefined_fields = BTreeSet::new();
    let Some(raw_paths) = object.get(PRESENCE_FIELD) else {
        return Ok(CapturedOperation {
            value: value.clone(),
            undefined_fields,
        });
    };
    let paths = raw_paths.as_array().ok_or_else(|| {
        WireError::structure(value, "Operation presence metadata is not an array.")
    })?;
    for raw_path in paths {
        let parts = raw_path
            .as_array()
            .filter(|parts| !parts.is_empty())
            .ok_or_else(|| {
                WireError::structure(
                    value,
                    "Operation presence metadata contains an invalid path.",
                )
            })?;
        let path = parts
            .iter()
            .map(|part| {
                let field = part
                    .as_str()
                    .filter(|field| !matches!(*field, "__proto__" | "constructor" | "prototype"))
                    .ok_or_else(|| {
                        WireError::structure(
                            value,
                            "Operation presence metadata contains an invalid path.",
                        )
                    })?;
                Ok(field.to_owned())
            })
            .collect::<Result<FieldPath, WireError>>()?;
        if matches!(path[0].as_str(), "type" | PRESENCE_FIELD) {
            return Err(WireError::structure(
                value,
                "Operation presence metadata cannot replace journal identity.",
            ));
        }
        if undefined_fields.contains(&path) {
            return Err(WireError::structure(
                value,
                "Operation presence metadata contains duplicate paths.",
            ));
        }
        let mut parent = value;
        let mut parent_path = Vec::new();
        for field in &path[..path.len() - 1] {
            parent_path.push(field.clone());
            if undefined_fields.contains(&parent_path) {
                return Err(WireError::structure(
                    value,
                    "Operation presence metadata has no object parent.",
                ));
            }
            parent = match parent {
                Value::Object(object) => object.get(field),
                Value::Array(array) => array_index(field).and_then(|index| array.get(index)),
                _ => None,
            }
            .ok_or_else(|| {
                WireError::structure(
                    value,
                    "Operation presence metadata has no owned parent field.",
                )
            })?;
        }
        let field = &path[path.len() - 1];
        match parent {
            Value::Object(object) => {
                if object.contains_key(field) {
                    return Err(WireError::structure(
                        value,
                        "Operation presence metadata conflicts with an existing value.",
                    ));
                }
            }
            Value::Array(array) => {
                let index = array_index(field)
                    .filter(|index| *index < array.len())
                    .ok_or_else(|| {
                        WireError::structure(
                            value,
                            "Operation presence metadata names an unavailable array slot.",
                        )
                    })?;
                if !array[index].is_null() {
                    return Err(WireError::structure(
                        value,
                        "Operation presence metadata conflicts with an existing value.",
                    ));
                }
            }
            _ => {
                return Err(WireError::structure(
                    value,
                    "Operation presence metadata has no object parent.",
                ));
            }
        }
        undefined_fields.insert(path);
    }
    Ok(CapturedOperation {
        value: value.clone(),
        undefined_fields,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn path(parts: &[&str]) -> FieldPath {
        parts.iter().map(|part| (*part).to_owned()).collect()
    }

    #[test]
    fn presence_roundtrip_preserves_all_four_states_and_unknown_fields() {
        let wire = json!({"schema":9,"op":{"type":"setRunProps","patch":{"bold":null,"italic":false},"unknown":{"rawXml":"<w:x/>"},"undefinedFields":[["patch","underline"]]}});
        let op = decode_envelope(&wire, 9).unwrap();
        assert_eq!(
            op.presence(&path(&["patch", "underline"])),
            FieldPresence::Undefined
        );
        assert_eq!(op.presence(&path(&["patch", "bold"])), FieldPresence::Null);
        assert_eq!(
            op.presence(&path(&["patch", "missing"])),
            FieldPresence::Absent
        );
        assert_eq!(
            op.presence(&path(&["patch", "italic"])),
            FieldPresence::Value(&json!(false))
        );
        assert_eq!(encode_envelope(&op, 9), wire);
        assert_eq!(op.patch_fields(&path(&["patch"])).unwrap().len(), 3);
    }

    #[test]
    fn presence_metadata_rejects_unsafe_conflicting_or_unowned_paths() {
        let cases = [
            json!(null),
            json!("invalid"),
            json!([[]]),
            json!([[1]]),
            json!([["type"]]),
            json!([["undefinedFields"]]),
            json!([["patch", "__proto__"]]),
            json!([["patch", "constructor"]]),
            json!([["patch", "prototype"]]),
            json!([["patch", "bold"]]),
            json!([["missing", "field"]]),
            json!([["patch", "x"], ["patch", "x"]]),
            json!([["slots", "01"]]),
            json!([["slots", "1"]]),
        ];
        for metadata in cases {
            let op = json!({"type":"setRunProps","patch":{"bold":null},"slots":[null],"undefinedFields":metadata});
            assert!(validate_operation(&op).is_err(), "accepted {op}");
        }
        let op = json!({"type":"insertContent","slots":[null],"undefinedFields":[["slots","0"]]});
        let captured = validate_operation(&op).unwrap();
        assert_eq!(
            captured.presence(&path(&["slots", "0"])),
            FieldPresence::Undefined
        );
    }

    #[test]
    fn envelope_uses_explicit_accepted_schema() {
        let wire = json!({"schema":10,"op":{"type":"future"}});
        assert_eq!(
            decode_envelope(&wire, 9).unwrap_err().reason,
            "unsupportedSchema"
        );
        assert!(decode_envelope(&wire, 10).is_ok());
    }
}
