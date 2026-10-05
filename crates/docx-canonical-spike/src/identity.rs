//! Structural identity census: derived section content is never counted twice.
use serde_json::Value;
use std::collections::BTreeMap;

fn visit(
    value: &Value,
    held_by: Option<&str>,
    owner_type: Option<&str>,
    keys: &mut Vec<String>,
    stacked: &mut BTreeMap<String, Vec<Value>>,
) {
    match value {
        Value::Array(items) => {
            for item in items {
                visit(item, held_by, owner_type, keys, stacked);
            }
        }
        Value::Object(fields) => {
            let revision = if held_by == Some("gridChange")
                && fields.get("columnWidths").is_some_and(Value::is_array)
            {
                fields.get("id")
            } else {
                fields
                    .get("info")
                    .filter(|info| info.get("author").is_some_and(Value::is_string))
                    .and_then(|info| info.get("id"))
            };
            if let Some(id) = revision.filter(|id| id.is_number()) {
                keys.push(format!("revision:{id}"));
            }
            let control = fields.get("id").filter(|id| id.is_number()).filter(|_| {
                (held_by == Some("properties")
                    && matches!(owner_type, Some("inlineSdt" | "blockSdt")))
                    || held_by == Some("contentControls")
                    || fields.get("sdtType").is_some_and(Value::is_string)
            });
            if let Some(id) = control {
                let key = format!("control:{id}");
                let same_stack = if held_by == Some("contentControls") {
                    let records = stacked.entry(key.clone()).or_default();
                    if records.contains(value) {
                        true
                    } else {
                        records.push(value.clone());
                        false
                    }
                } else {
                    false
                };
                if !same_stack {
                    keys.push(key);
                }
            }
            let kind = fields.get("type").and_then(Value::as_str);
            for (key, child) in fields {
                visit(child, Some(key), kind, keys, stacked);
            }
        }
        _ => {}
    }
}

pub(crate) fn keys(value: &Value) -> Vec<String> {
    let mut keys = Vec::new();
    visit(value, None, None, &mut keys, &mut BTreeMap::new());
    keys
}

pub(crate) fn package_keys(document: &Value) -> Vec<String> {
    let mut package = document["package"].clone();
    if let Some(body) = package.get_mut("document").and_then(Value::as_object_mut) {
        body.remove("sections");
    }
    keys(&package)
}
