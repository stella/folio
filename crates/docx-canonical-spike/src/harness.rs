//! Test-only tagged transport. This is not a document or operation wire schema.
//! Special JS leaves stay in path sidecars. Edits that would relocate or capture
//! those leaves into inverses are explicitly unsupported until relocation exists.

use crate::apply::{Applied, Touched};
use crate::comments::RelationshipState;
use crate::refusal::Failure;
use serde_json::{Map, Value, json};
use std::collections::BTreeMap;

type Path = Vec<String>;

#[derive(Clone, Debug, PartialEq)]
enum Special {
    Map,
    NullObject,
    Date,
    Undefined,
    Uint8Array,
    ArrayBuffer,
    Hole,
}

#[derive(Clone, Debug)]
pub struct Decoded {
    pub value: Value,
    specials: BTreeMap<Path, Special>,
    object_order: BTreeMap<Path, Vec<String>>,
    map_order: BTreeMap<Path, Vec<String>>,
}

fn fields<'a>(
    value: &'a Value,
    tag: &str,
    keys: &[&str],
) -> Result<&'a Map<String, Value>, String> {
    let object = value
        .as_object()
        .ok_or_else(|| format!("Expected a {tag} tag object."))?;
    if object.len() != keys.len() + 1
        || object
            .keys()
            .any(|key| key != "tag" && !keys.contains(&key.as_str()))
    {
        return Err(format!("Unexpected fields in the {tag} tag."));
    }
    Ok(object)
}

/// The canonical spelling emitted by JS Date.toISOString, including extended
/// years and the finite Date range. No timezone or calendar normalization.
fn valid_iso(iso: &str) -> bool {
    if !iso.is_ascii() {
        return false;
    }
    let extended = iso.starts_with('+') || iso.starts_with('-');
    let year_length = if extended { 7 } else { 4 };
    if iso.len() != year_length + 20 {
        return false;
    }
    let year_digits = if extended { &iso[1..7] } else { &iso[..4] };
    if !year_digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    let Some(mut year) = year_digits.parse::<i64>().ok() else {
        return false;
    };
    if extended && (iso.starts_with('+') && year < 10_000 || iso.starts_with('-') && year == 0) {
        return false;
    }
    if iso.starts_with('-') {
        year = -year;
    }
    let rest = &iso[year_length..];
    for (index, separator) in [
        (0, b'-'),
        (3, b'-'),
        (6, b'T'),
        (9, b':'),
        (12, b':'),
        (15, b'.'),
        (19, b'Z'),
    ] {
        if rest.as_bytes()[index] != separator {
            return false;
        }
    }
    let number = |from, to| {
        let digits = &rest[from..to];
        digits
            .bytes()
            .all(|byte| byte.is_ascii_digit())
            .then(|| digits.parse::<i64>().ok())
            .flatten()
    };
    let (Some(month), Some(day), Some(hour), Some(minute), Some(second), Some(millis)) = (
        number(1, 3),
        number(4, 6),
        number(7, 9),
        number(10, 12),
        number(13, 15),
        number(16, 19),
    ) else {
        return false;
    };
    if !(1..=12).contains(&month) || hour > 23 || minute > 59 || second > 59 {
        return false;
    }
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days = match month {
        2 => {
            if leap {
                29
            } else {
                28
            }
        }
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    };
    if day < 1 || day > days {
        return false;
    }
    let adjusted_year = year - if month <= 2 { 1 } else { 0 };
    let era = (if adjusted_year >= 0 {
        adjusted_year
    } else {
        adjusted_year - 399
    }) / 400;
    let year_of_era = adjusted_year - era * 400;
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + day - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    let epoch_days = era * 146_097 + day_of_era - 719_468;
    let epoch_millis =
        epoch_days * 86_400_000 + hour * 3_600_000 + minute * 60_000 + second * 1000 + millis;
    (-8_640_000_000_000_000..=8_640_000_000_000_000).contains(&epoch_millis)
}

fn decode_node(
    value: &Value,
    path: &mut Path,
    decoded: &mut Decoded,
    slot: bool,
) -> Result<Option<Value>, String> {
    if !value.is_object() {
        if value.is_array() {
            return Err("Raw arrays are not tagged harness values.".into());
        }
        if value
            .as_f64()
            .is_some_and(|number| number == 0.0 && number.is_sign_negative())
        {
            return Err("Negative zero is outside the harness codec.".into());
        }
        return Ok(Some(value.clone()));
    }
    let tag = value["tag"]
        .as_str()
        .ok_or("A harness object needs a tag.")?;
    match tag {
        "object" | "nullObject" | "map" => {
            fields(value, tag, &["entries"])?;
            let entries = value["entries"]
                .as_array()
                .ok_or("Object and Map entries must be arrays.")?;
            let mut object = Map::new();
            let mut order = Vec::new();
            for entry in entries {
                let pair = entry
                    .as_array()
                    .filter(|pair| pair.len() == 2)
                    .ok_or("An entry must be a key/value pair.")?;
                let key = pair[0]
                    .as_str()
                    .ok_or("Harness object and Map keys must be strings.")?;
                if order.iter().any(|old| old == key) {
                    return Err("Harness entries repeat an owned key.".into());
                }
                order.push(key.to_owned());
                path.push(key.to_owned());
                let child = decode_node(&pair[1], path, decoded, false)?;
                path.pop();
                if let Some(child) = child {
                    object.insert(key.to_owned(), child);
                }
            }
            if tag == "map" {
                decoded.map_order.insert(path.clone(), order);
                decoded.specials.insert(path.clone(), Special::Map);
            } else {
                decoded.object_order.insert(path.clone(), order);
                if tag == "nullObject" {
                    decoded.specials.insert(path.clone(), Special::NullObject);
                }
            }
            Ok(Some(Value::Object(object)))
        }
        "array" => {
            fields(value, tag, &["items"])?;
            let items = value["items"]
                .as_array()
                .ok_or("Array items must be an array.")?;
            let mut array = Vec::new();
            for (index, item) in items.iter().enumerate() {
                path.push(index.to_string());
                array.push(decode_node(item, path, decoded, true)?.unwrap_or(Value::Null));
                path.pop();
            }
            Ok(Some(Value::Array(array)))
        }
        "date" => {
            fields(value, tag, &["iso"])?;
            let iso = value["iso"].as_str().ok_or("Date ISO must be a string.")?;
            if !valid_iso(iso) {
                return Err("Invalid ISO Date.".into());
            }
            decoded.specials.insert(path.clone(), Special::Date);
            Ok(Some(Value::String(iso.to_owned())))
        }
        "undefined" | "hole" => {
            fields(value, tag, &[])?;
            if tag == "hole" && !slot {
                return Err("A hole tag belongs only to an array slot.".into());
            }
            decoded.specials.insert(
                path.clone(),
                if tag == "hole" {
                    Special::Hole
                } else {
                    Special::Undefined
                },
            );
            Ok(None)
        }
        "uint8Array" | "arrayBuffer" => {
            fields(value, tag, &["bytes"])?;
            let bytes = value["bytes"]
                .as_array()
                .ok_or("Binary bytes must be an array.")?;
            if bytes
                .iter()
                .any(|byte| byte.as_u64().is_none_or(|byte| byte > 255))
            {
                return Err("Binary bytes must be integers in 0..255.".into());
            }
            decoded.specials.insert(
                path.clone(),
                if tag == "uint8Array" {
                    Special::Uint8Array
                } else {
                    Special::ArrayBuffer
                },
            );
            Ok(Some(Value::Array(bytes.clone())))
        }
        _ => Err(format!("Unknown harness tag {tag}.")),
    }
}

pub fn decode(value: &Value) -> Result<Decoded, String> {
    let mut decoded = Decoded {
        value: Value::Null,
        specials: BTreeMap::new(),
        object_order: BTreeMap::new(),
        map_order: BTreeMap::new(),
    };
    decoded.value = decode_node(value, &mut vec![], &mut decoded, false)?.unwrap_or(Value::Null);
    Ok(decoded)
}

fn at<'a>(value: &'a Value, path: &[String]) -> Option<&'a Value> {
    let mut value = value;
    for part in path {
        value = match value {
            Value::Array(items) => items.get(part.parse::<usize>().ok()?)?,
            Value::Object(fields) => fields.get(part)?,
            _ => return None,
        };
    }
    Some(value)
}

fn encode_node(value: &Value, path: &mut Path, decoded: &Decoded) -> Result<Value, String> {
    match decoded.specials.get(path) {
        Some(Special::Undefined) => return Ok(json!({"tag":"undefined"})),
        Some(Special::Hole) => return Ok(json!({"tag":"hole"})),
        Some(Special::Date) => {
            let iso = value
                .as_str()
                .ok_or("A Date sidecar no longer points to an ISO string.")?;
            if !valid_iso(iso) {
                return Err("Invalid ISO Date.".into());
            }
            return Ok(json!({"tag":"date","iso":iso}));
        }
        Some(Special::Uint8Array | Special::ArrayBuffer) => {
            let bytes = value
                .as_array()
                .ok_or("A binary sidecar no longer points to bytes.")?;
            if bytes
                .iter()
                .any(|byte| byte.as_u64().is_none_or(|byte| byte > 255))
            {
                return Err("A binary sidecar has invalid bytes.".into());
            }
            return Ok(
                json!({"tag":if decoded.specials.get(path)==Some(&Special::Uint8Array){"uint8Array"}else{"arrayBuffer"},"bytes":bytes}),
            );
        }
        _ => {}
    }
    match value {
        Value::Object(fields) => {
            let is_map = decoded.specials.get(path) == Some(&Special::Map);
            let mut keys = if is_map {
                decoded
                    .map_order
                    .get(path)
                    .cloned()
                    .ok_or("A Map sidecar has no key order.")?
            } else {
                decoded.object_order.get(path).cloned().unwrap_or_default()
            };
            for key in fields.keys() {
                if !keys.contains(key) {
                    keys.push(key.clone());
                }
            }
            let mut entries = Vec::new();
            for key in keys {
                path.push(key.clone());
                let child = fields.get(&key);
                let special = decoded.specials.get(path);
                if let Some(child) = child {
                    entries.push(json!([key, encode_node(child, path, decoded)?]));
                } else if special == Some(&Special::Undefined) {
                    entries.push(json!([key,{"tag":"undefined"}]));
                }
                path.pop();
            }
            let tag = if is_map {
                "map"
            } else if decoded.specials.get(path) == Some(&Special::NullObject) {
                "nullObject"
            } else {
                "object"
            };
            Ok(json!({"tag":tag,"entries":entries}))
        }
        Value::Array(items) => {
            let mut encoded = Vec::new();
            for (index, item) in items.iter().enumerate() {
                path.push(index.to_string());
                encoded.push(encode_node(item, path, decoded)?);
                path.pop();
            }
            Ok(json!({"tag":"array","items":encoded}))
        }
        _ => {
            if decoded.specials.get(path) == Some(&Special::Map) {
                return Err("A Map sidecar no longer points to an object.".into());
            }
            if decoded.specials.get(path) == Some(&Special::NullObject) {
                return Err("A null-prototype sidecar no longer points to an object.".into());
            }
            Ok(value.clone())
        }
    }
}

pub fn encode(decoded: &Decoded) -> Result<Value, String> {
    encode_node(&decoded.value, &mut vec![], decoded)
}
pub fn round_trip(encoded: &Value) -> Result<Value, String> {
    encode(&decode(encoded)?)
}

fn relationship_path() -> Path {
    vec!["package".into(), "relationships".into()]
}
fn comment_kind(op: &Value) -> bool {
    matches!(
        op["type"].as_str(),
        Some("createComment" | "deleteComment" | "restoreCommentState")
    )
}
fn target_ids(op: &Value) -> Vec<&str> {
    [
        op.get("blockId"),
        op.pointer("/at/blockId"),
        op.pointer("/from/blockId"),
        op.pointer("/to/blockId"),
        op.get("first"),
        op.get("second"),
    ]
    .into_iter()
    .flatten()
    .filter_map(Value::as_str)
    .collect()
}

fn paragraph_owner<'a>(document: &'a Value, path: &[String]) -> Option<(&'a str, usize)> {
    for length in 0..=path.len() {
        let value = at(document, &path[..length])?;
        if value["type"] == "paragraph" {
            return value["paraId"].as_str().map(|id| (id, length));
        }
    }
    None
}

fn guard(decoded: &Decoded, op: &Value) -> Result<(), Failure> {
    let structural = matches!(
        op["type"].as_str(),
        Some(
            "splitBlock"
                | "joinBlocks"
                | "insertBlocks"
                | "deleteBlocks"
                | "replaceBlocks"
                | "insertTable"
                | "deleteTable"
                | "setContainerBlocks"
                | "insertRow"
                | "deleteRow"
                | "setTableRows"
                | "setTable"
        )
    );
    let ids = target_ids(op);
    for (path, special) in &decoded.specials {
        if path == &relationship_path() && comment_kind(op) {
            continue;
        }
        if comment_kind(op)
            && path.starts_with(&vec![
                "package".into(),
                "document".into(),
                "comments".into(),
            ])
        {
            return Err(Failure::unsupported(op, "harnessSidecarMovement"));
        }
        if comment_kind(op) && path.starts_with(&relationship_path()) && path.len() > 2 {
            let relation_path = &path[..3];
            let relation = at(&decoded.value, relation_path);
            if relation.is_some_and(|relation|matches!(relation["type"].as_str(),Some(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"|
                "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"))) {
                return Err(Failure::unsupported(op,"harnessSidecarMovement"));
            }
        }
        let in_body =
            path.starts_with(&vec!["package".into(), "document".into(), "content".into()])
                || path.starts_with(&vec![
                    "package".into(),
                    "document".into(),
                    "sections".into(),
                ]);
        if structural && in_body {
            return Err(Failure::unsupported(op, "harnessSidecarMovement"));
        }
        if let Some((id, owner_length)) = paragraph_owner(&decoded.value, path) {
            let content_field = path.get(owner_length).is_some_and(|key| key == "content");
            let formatting_field = path.get(owner_length).is_some_and(|key| {
                matches!(key.as_str(), "formatting" | "propertyChanges" | "pPrMark")
            });
            let targeted = ids.iter().any(|target| target.eq_ignore_ascii_case(id));
            if targeted
                && (content_field
                    || formatting_field
                        && matches!(
                            op["type"].as_str(),
                            Some("setParagraphProps" | "setParagraphReview")
                        ))
            {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
            // Comment anchors discover owned paragraphs across the document.
            if comment_kind(op) && content_field {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
        }
        if matches!(special, Special::Undefined | Special::Hole) && path.is_empty() {
            return Err(Failure::unsupported(op, "harnessRootPresence"));
        }
    }
    Ok(())
}

fn comment_edit(decoded: &mut Decoded, op: &Value) -> Result<Applied, Failure> {
    let path = relationship_path();
    let previous = match decoded.specials.get(&path) {
        Some(Special::Map) => {
            let fields = at(&decoded.value, &path)
                .and_then(Value::as_object)
                .ok_or_else(|| Failure::unsupported(op, "documentMapTransport"))?;
            // Own undefined Map values cannot yet be carried through relationship captures.
            if decoded.specials.iter().any(|(child, special)| {
                child.len() == 3 && child.starts_with(&path) && *special == Special::Undefined
            }) {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
            let keys = decoded
                .map_order
                .get(&path)
                .ok_or_else(|| Failure::unsupported(op, "documentMapTransport"))?;
            if keys.len() != fields.len() || keys.iter().any(|key| !fields.contains_key(key)) {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
            RelationshipState::Present(
                keys.iter()
                    .map(|key| (key.clone(), fields[key].clone()))
                    .collect(),
            )
        }
        Some(Special::Undefined) => RelationshipState::Undefined,
        None if at(&decoded.value, &path).is_none() => RelationshipState::Absent,
        _ => return Err(Failure::unsupported(op, "documentMapTransport")),
    };
    let mut document = decoded.value.clone();
    let package = document["package"]
        .as_object_mut()
        .ok_or_else(|| Failure::unsupported(op, "documentShape"))?;
    package.remove("relationships");
    let mut result = crate::comments::edit_with_relationships(&document, op, previous)?;
    crate::apply::refresh_sections(&mut result.applied.document, op)?;
    decoded.specials.remove(&path);
    decoded.map_order.remove(&path);
    let package = result.applied.document["package"]
        .as_object_mut()
        .expect("comment document package");
    match result.relationships {
        RelationshipState::Absent => {}
        RelationshipState::Undefined => {
            decoded.specials.insert(path, Special::Undefined);
        }
        RelationshipState::Present(entries) => {
            decoded.map_order.insert(
                path.clone(),
                entries.iter().map(|(key, _)| key.clone()).collect(),
            );
            package.insert(
                "relationships".into(),
                Value::Object(entries.into_iter().collect()),
            );
            decoded.specials.insert(path, Special::Map);
        }
    }
    Ok(result.applied)
}

#[derive(Clone, Copy, PartialEq)]
enum Touch {
    Modified,
    Inserted,
    Removed,
}
fn combine(previous: Option<Touch>, next: Touch) -> Option<Touch> {
    use Touch::*;
    match (previous, next) {
        (None, next) => Some(next),
        (Some(Modified), Modified | Inserted) => Some(Modified),
        (Some(Modified), Removed) => Some(Removed),
        (Some(Inserted), Modified | Inserted) => Some(Inserted),
        (Some(Inserted), Removed) => None,
        (Some(Removed), Modified | Inserted) => Some(Modified),
        (Some(Removed), Removed) => Some(Removed),
    }
}

fn apply_decoded(mut decoded: Decoded, operations: &[Value]) -> Result<Decoded, Failure> {
    let mut inverse = Vec::new();
    let mut touched = BTreeMap::new();
    let mut order = Vec::new();
    let mut revisions = Vec::new();
    for op in operations {
        crate::wire::validate_operation(op)
            .map_err(|error| Failure::refused(op, error.reason, error.message))?;
        guard(&decoded, op)?;
        let applied = if comment_kind(op) {
            comment_edit(&mut decoded, op)?
        } else {
            crate::apply::apply_operations(&decoded.value, std::slice::from_ref(op))?
        };
        for (path, special) in &decoded.specials {
            if comment_kind(op) && path == &relationship_path() {
                continue;
            }
            if matches!(special, Special::Undefined | Special::Hole) {
                let parent = path.split_last().map(|(_, parent)| parent).unwrap_or(&[]);
                let before = at(&decoded.value, parent);
                let after = at(&applied.document, parent);
                if before.is_some_and(Value::is_array) && before != after
                    || before.is_some_and(Value::is_object) && !after.is_some_and(Value::is_object)
                    || *special == Special::Undefined
                        && at(&decoded.value, path).is_none()
                        && at(&applied.document, path).is_some()
                {
                    return Err(Failure::unsupported(op, "harnessSidecarMovement"));
                }
                continue;
            }
            if at(&decoded.value, path) != at(&applied.document, path) && *special != Special::Map {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
            if *special == Special::Map && at(&applied.document, path).is_none() {
                return Err(Failure::unsupported(op, "harnessSidecarMovement"));
            }
        }
        decoded.value = applied.document;
        let mut next_inverse = applied.inverse;
        next_inverse.extend(inverse);
        inverse = next_inverse;
        for (state, ids) in [
            (Touch::Modified, applied.touched.modified),
            (Touch::Inserted, applied.touched.inserted),
            (Touch::Removed, applied.touched.removed),
        ] {
            for id in ids {
                let previous = touched.get(&id).copied();
                match combine(previous, state) {
                    Some(next) => {
                        if previous.is_none() {
                            order.push(id.clone());
                        }
                        touched.insert(id, next);
                    }
                    None => {
                        touched.remove(&id);
                        order.retain(|previous| previous != &id);
                    }
                }
            }
        }
        revisions.extend(applied.revisions);
    }
    let ids = |state| {
        order
            .iter()
            .filter(|id| touched.get(*id) == Some(&state))
            .cloned()
            .collect::<Vec<_>>()
    };
    let applied = Applied {
        document: decoded.value,
        inverse,
        touched: Touched {
            modified: ids(Touch::Modified),
            inserted: ids(Touch::Inserted),
            removed: ids(Touch::Removed),
        },
        revisions,
    };
    let specials = decoded
        .specials
        .into_iter()
        .map(|(mut path, special)| {
            path.insert(0, "document".into());
            (path, special)
        })
        .collect();
    let object_order = decoded
        .object_order
        .into_iter()
        .map(|(mut path, order)| {
            path.insert(0, "document".into());
            (path, order)
        })
        .collect();
    let map_order = decoded
        .map_order
        .into_iter()
        .map(|(mut path, order)| {
            path.insert(0, "document".into());
            (path, order)
        })
        .collect();
    Ok(Decoded {
        value: serde_json::to_value(applied).expect("Applied is JSON-backed"),
        specials,
        object_order,
        map_order,
    })
}

/// Receives the inner `harness` request. Failures are plain; success is tagged.
pub fn apply_request(request: &Value) -> Value {
    if let Some(encoded) = request.get("roundtrip") {
        return match round_trip(encoded) {
            Ok(encoded) => json!({"harness":encoded}),
            Err(message) => json!({"status":"transportError","message":message}),
        };
    }
    let Some(operations) = request["ops"].as_array() else {
        return json!({"status":"transportError","message":"Expected harness document and ops array."});
    };
    let Some(document) = request.get("document") else {
        return json!({"status":"transportError","message":"Expected harness document and ops array."});
    };
    let decoded = match decode(document) {
        Ok(decoded) => decoded,
        Err(message) => return json!({"status":"transportError","message":message}),
    };
    match apply_decoded(decoded, operations) {
        Err(failure) => serde_json::to_value(failure).expect("Failure is serializable"),
        Ok(decoded) => match encode(&decoded) {
            Ok(encoded) => json!({"harness":encoded}),
            Err(message) => json!({"status":"transportError","message":message}),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn encode_plain(value: &Value) -> Value {
        encode(&Decoded {
            value: value.clone(),
            specials: BTreeMap::new(),
            object_order: BTreeMap::new(),
            map_order: BTreeMap::new(),
        })
        .unwrap()
    }
    #[test]
    fn negative_zero_refuses_instead_of_losing_its_sign() {
        let negative_zero: Value = serde_json::from_str("-0.0").unwrap();
        assert!(decode(&negative_zero).is_err());
        assert_eq!(round_trip(&json!(0)).unwrap(), json!(0));
    }
    #[test]
    fn all_tags_and_owned_undefined_round_trip_without_json_equivalence() {
        let value = json!({"tag":"object","entries":[["missing",{"tag":"undefined"}],["null",null],
            ["date",{"tag":"date","iso":"2026-10-05T00:00:00.000Z"}],
            ["map",{"tag":"map","entries":[["z",{"tag":"undefined"}],["a",{"tag":"date","iso":"2025-01-01T00:00:00.000Z"}]]}],
            ["array",{"tag":"array","items":[{"tag":"hole"},{"tag":"undefined"},null]}],
            ["bytes",{"tag":"uint8Array","bytes":[0,127,255]}],["buffer",{"tag":"arrayBuffer","bytes":[]}]]});
        let decoded = decode(&value).unwrap();
        assert!(decoded.value.get("missing").is_none());
        assert_eq!(decoded.value["null"], Value::Null);
        assert_eq!(encode(&decoded).unwrap(), value);
    }
    #[test]
    fn null_prototype_records_preserve_tag_special_keys_and_owned_undefined() {
        let value = json!({"tag":"nullObject","entries":[
            ["__proto__",{"tag":"nullObject","entries":[]}],
            ["constructor",{"tag":"undefined"}],
            ["toString",null],
            ["map",{"tag":"map","entries":[["nested",{"tag":"nullObject","entries":[]}]]}]
        ]});
        let decoded = decode(&value).unwrap();
        assert_eq!(decoded.specials.get(&vec![]), Some(&Special::NullObject));
        assert!(decoded.value.get("constructor").is_none());
        assert_eq!(encode(&decoded).unwrap(), value);
    }
    #[test]
    fn every_undefined_map_entry_mask_preserves_exact_iteration_order() {
        let keys = ["z", "5", "a", "2", "m"];
        for mask in 0..(1 << keys.len()) {
            let entries: Vec<_> = keys
                .iter()
                .enumerate()
                .map(|(index, key)| {
                    json!([
                        key,
                        if mask & (1 << index) == 0 {
                            json!(index)
                        } else {
                            json!({"tag":"undefined"})
                        }
                    ])
                })
                .collect();
            let tagged = json!({"tag":"map","entries":entries});
            let decoded = decode(&tagged).unwrap();
            assert_eq!(decoded.map_order[&vec![]], keys.map(str::to_owned));
            assert_eq!(encode(&decoded).unwrap(), tagged, "undefined mask {mask}");
            assert_eq!(
                apply_request(&json!({"roundtrip":tagged}))["harness"],
                tagged
            );
        }
    }
    #[test]
    fn map_order_sidecar_uses_current_values_and_supported_new_keys() {
        let tagged = json!({"tag":"map","entries":[["z",0],["missing",{"tag":"undefined"}],["a",1],["removed",2]]});
        let mut decoded = decode(&tagged).unwrap();
        decoded.value = json!({"a":3,"newB":4,"z":5,"newA":6});
        assert_eq!(
            encode(&decoded).unwrap(),
            json!({"tag":"map","entries":[["z",5],["missing",{"tag":"undefined"}],["a",3],["newB",4],["newA",6]]})
        );
    }
    #[test]
    fn comment_relationship_recomputation_owns_the_returned_map_order() {
        let document = json!({"package":{"document":{"content":[{"type":"paragraph","paraId":"00000001","content":[]}]},
            "relationships":{
                "rId8":{"id":"rId8","type":"http://schemas.microsoft.com/office/2011/relationships/commentsExtended","target":"commentsExtended.xml"},
                "rId3":{"id":"rId3","type":"unowned","target":"keep"},
                "rId5":{"id":"rId5","type":"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments","target":"comments.xml"}}}});
        let mut decoded = decode(&encode_plain(&document)).unwrap();
        let path = relationship_path();
        decoded.object_order.remove(&path);
        decoded.specials.insert(path.clone(), Special::Map);
        decoded.map_order.insert(
            path.clone(),
            vec!["rId8".into(), "rId3".into(), "rId5".into()],
        );
        let op = json!({"type":"createComment","comment":{"id":3,"author":"author","content":[{"type":"paragraph","paraId":"00000003","content":[]}]},
            "anchor":{"kind":"point","at":{"story":"main","blockId":"00000001","offset":0}}});
        let created = comment_edit(&mut decoded, &op).unwrap();
        assert_eq!(decoded.map_order[&path], vec!["rId3", "rId5"]);
        decoded.value = created.document;
        let mut restore = created.inverse[0].clone();
        restore["state"]["relationships"][0]["index"] = json!(1);
        let restored = comment_edit(&mut decoded, &restore).unwrap();
        assert_eq!(decoded.map_order[&path], vec!["rId3", "rId8", "rId5"]);
        decoded.value = restored.document;
        let reencoded = decode(&encode(&decoded).unwrap()).unwrap();
        assert_eq!(reencoded.map_order[&path], vec!["rId3", "rId8", "rId5"]);
    }
    #[test]
    fn iso_dates_match_js_canonical_calendar_and_range_boundaries() {
        for iso in [
            "0000-02-29T00:00:00.000Z",
            "2000-02-29T23:59:59.999Z",
            "-000001-01-01T00:00:00.000Z",
            "+010000-01-01T00:00:00.000Z",
            "-271821-04-20T00:00:00.000Z",
            "+275760-09-13T00:00:00.000Z",
        ] {
            assert!(valid_iso(iso), "{iso}");
        }
        for iso in [
            "1900-02-29T00:00:00.000Z",
            "2026-01-01T24:00:00.000Z",
            "2026-01-01T00:00:60.000Z",
            "+000001-01-01T00:00:00.000Z",
            "-000000-01-01T00:00:00.000Z",
            "+275760-09-13T00:00:00.001Z",
            "-271821-04-19T23:59:59.999Z",
        ] {
            assert!(!valid_iso(iso), "{iso}");
        }
    }
    #[test]
    fn tagged_success_keeps_package_leaves_and_refuses_moving_text_sidecars() {
        let mut document = json!({"package":{"document":{"content":[{"type":"paragraph","paraId":"00000001","content":[
            {"type":"run","content":[{"type":"text","text":"abc"}]}]}]},"metadata":"2026-10-05T00:00:00.000Z"}});
        let mut decoded = decode(&encode_plain(&document)).unwrap();
        decoded
            .specials
            .insert(vec!["package".into(), "metadata".into()], Special::Date);
        let op = json!({"type":"insertText","at":{"story":"main","blockId":"00000001","offset":1},"text":"x","runProps":"inherit"});
        let changed = apply_decoded(decoded.clone(), std::slice::from_ref(&op)).unwrap();
        assert_eq!(
            changed.specials[&vec!["document".into(), "package".into(), "metadata".into()]],
            Special::Date
        );
        document["package"]["document"]["content"][0]["content"][0]["metadata"] = json!("date");
        decoded.value = document;
        decoded.specials.insert(
            vec![
                "package".into(),
                "document".into(),
                "content".into(),
                "0".into(),
                "content".into(),
                "0".into(),
                "metadata".into(),
            ],
            Special::Date,
        );
        assert!(
            matches!(apply_decoded(decoded,&[op]),Err(Failure::Unsupported{dimension,..}) if dimension=="harnessSidecarMovement")
        );
    }
    #[test]
    fn map_relationship_comments_encode_real_map_without_document_codec_changes() {
        let document = json!({"package":{"document":{"content":[{"type":"paragraph","paraId":"00000001","content":[]}]}}});
        let op = json!({"type":"createComment","comment":{"id":3,"author":"author","content":[{"type":"paragraph","paraId":"00000003","content":[]}]},
            "anchor":{"kind":"point","at":{"story":"main","blockId":"00000001","offset":0}}});
        let result = apply_request(&json!({"document":encode_plain(&document),"ops":[op]}));
        let encoded = &result["harness"];
        assert!(encoded.is_object(), "{result}");
        let decoded = decode(encoded).unwrap();
        assert_eq!(
            decoded.specials[&vec!["document".into(), "package".into(), "relationships".into()]],
            Special::Map
        );
        let inverse = decoded.value["inverse"].clone();
        let mut document_encoded = decoded.clone();
        document_encoded.value = decoded.value["document"].clone();
        document_encoded.specials = decoded
            .specials
            .into_iter()
            .filter_map(|(path, special)| {
                path.strip_prefix(&["document".into()])
                    .map(|path| (path.to_vec(), special))
            })
            .collect();
        document_encoded.object_order = decoded
            .object_order
            .into_iter()
            .filter_map(|(path, order)| {
                path.strip_prefix(&["document".into()])
                    .map(|path| (path.to_vec(), order))
            })
            .collect();
        document_encoded.map_order = decoded
            .map_order
            .into_iter()
            .filter_map(|(path, order)| {
                path.strip_prefix(&["document".into()])
                    .map(|path| (path.to_vec(), order))
            })
            .collect();
        let restored =
            apply_request(&json!({"document":encode(&document_encoded).unwrap(),"ops":inverse}));
        let restored = decode(&restored["harness"]).unwrap();
        assert_eq!(restored.value["document"], document);
        assert!(!restored.specials.contains_key(&vec![
            "document".into(),
            "package".into(),
            "relationships".into()
        ]));
    }
}
