//! Atomic operations over the lossless model. Unsupported dimensions are explicit.
use serde::Serialize;
use serde_json::{Map, Value, json};
use std::collections::BTreeMap;

use crate::refusal::Failure;

#[derive(Debug, Clone, Default, Serialize)]
pub struct Touched {
    pub modified: Vec<String>,
    pub inserted: Vec<String>,
    pub removed: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Applied {
    pub document: Value,
    pub inverse: Vec<Value>,
    pub touched: Touched,
    pub revisions: Vec<i64>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Step {
    Field(String),
    Index(usize),
}

pub(crate) type Path = Vec<Step>;

pub(crate) fn at_path<'a>(root: &'a Value, path: &[Step]) -> Option<&'a Value> {
    let mut current = root;
    for step in path {
        current = match step {
            Step::Field(key) => current.get(key)?,
            Step::Index(index) => current.get(*index)?,
        };
    }
    Some(current)
}

pub(crate) fn at_path_mut<'a>(root: &'a mut Value, path: &[Step]) -> Option<&'a mut Value> {
    let mut current = root;
    for step in path {
        current = match step {
            Step::Field(key) => current.get_mut(key)?,
            Step::Index(index) => current.get_mut(*index)?,
        };
    }
    Some(current)
}

pub(crate) fn paragraph_paths(value: &Value, path: &mut Path, out: &mut Vec<Path>) {
    if value.get("type").and_then(Value::as_str) == Some("paragraph") {
        out.push(path.clone());
        return;
    }
    match value {
        Value::Array(items) => {
            for (index, item) in items.iter().enumerate() {
                path.push(Step::Index(index));
                paragraph_paths(item, path, out);
                path.pop();
            }
        }
        Value::Object(fields) => {
            // Only container edges; captured JSON elsewhere is not a paragraph owner.
            for key in ["content", "rows", "cells"] {
                if let Some(child) = fields.get(key) {
                    path.push(Step::Field(key.to_owned()));
                    paragraph_paths(child, path, out);
                    path.pop();
                }
            }
        }
        _ => {}
    }
}

pub(crate) fn locate(document: &Value, op: &Value) -> Result<Path, Failure> {
    if op.get("story").and_then(Value::as_str) != Some("main") {
        return Err(Failure::unsupported(op, "secondaryStory"));
    }
    let block_id = required_string(op, "blockId")?;
    let body = document
        .pointer("/package/document")
        .ok_or_else(|| Failure::unsupported(op, "documentShape"))?;
    let mut paths = Vec::new();
    paragraph_paths(body, &mut vec![], &mut paths);
    let found = paths.into_iter().find(|path| {
        at_path(body, path)
            .and_then(|p| p.get("paraId"))
            .and_then(Value::as_str)
            .is_some_and(|id| id.eq_ignore_ascii_case(block_id))
    });
    let path = found.ok_or_else(|| {
        Failure::refused(op, "blockNotFound", format!("No paragraph is {block_id}."))
    })?;
    let mut full = vec![
        Step::Field("package".into()),
        Step::Field("document".into()),
    ];
    full.extend(path);
    Ok(full)
}

pub(crate) fn required_string<'a>(op: &'a Value, key: &str) -> Result<&'a str, Failure> {
    op.get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| Failure::refused(op, "invalidOperation", format!("Operation has no {key}.")))
}

fn review_fields(paragraph: &Value) -> Value {
    let mut fields = Map::new();
    for key in ["formatting", "propertyChanges", "pPrMark"] {
        if let Some(value) = paragraph.get(key) {
            fields.insert(key.into(), value.clone());
        }
    }
    Value::Object(fields)
}

fn states_values(base: Option<&Value>, expected: &Map<String, Value>) -> bool {
    expected.iter().all(|(key, value)| {
        if value.is_null() {
            base.and_then(|base| base.get(key)).is_none()
        } else {
            base.and_then(|base| base.get(key)) == Some(value)
        }
    })
}

fn set_paragraph_props(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    let patch = op.get("patch").and_then(Value::as_object).ok_or_else(|| {
        Failure::refused(op, "invalidOperation", "Operation has no property patch.")
    })?;
    if op.get("revision").is_some()
        && ["runProperties", "runInWithNext"]
            .iter()
            .any(|key| patch.contains_key(*key))
    {
        return Err(Failure::refused(
            op,
            "untrackable",
            "A paragraph property change does not record the paragraph mark's run properties.",
        ));
    }
    let previous = paragraph.get("formatting");
    if let Some(expected) = op.get("expected").and_then(Value::as_object) {
        if !states_values(previous, expected) {
            return Err(Failure::refused(
                op,
                "stale",
                format!(
                    "{} states other paragraph properties than expected.",
                    required_string(op, "blockId")?
                ),
            ));
        }
    }
    let mut formatting = previous
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let mut restore = Map::new();
    for (key, value) in patch {
        restore.insert(
            key.clone(),
            previous
                .and_then(|base| base.get(key))
                .cloned()
                .unwrap_or(Value::Null),
        );
        if value.is_null() {
            formatting.remove(key);
        } else {
            formatting.insert(key.clone(), value.clone());
        }
    }
    let mut result = paragraph.clone();
    let fields = result
        .as_object_mut()
        .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?;
    if formatting.is_empty() && op.get("whenEmpty").and_then(Value::as_str) != Some("keep") {
        fields.remove("formatting");
    } else {
        fields.insert("formatting".into(), Value::Object(formatting));
    }
    if result == *paragraph {
        return Ok((result, vec![]));
    }
    if let Some(stamp) = op.get("revision") {
        let existing = paragraph
            .get("propertyChanges")
            .and_then(Value::as_array)
            .and_then(|changes| changes.first());
        let mut info = existing
            .and_then(|change| change.get("info"))
            .and_then(Value::as_object)
            .cloned()
            .unwrap_or_default();
        for key in ["id", "author", "date"] {
            if let Some(value) = stamp.get(key) {
                info.insert(key.into(), value.clone());
            }
        }
        if let Some(id) = existing
            .and_then(|change| change.get("info"))
            .and_then(|info| info.get("id"))
        {
            info.insert("id".into(), id.clone());
        }
        info.remove("initials");
        if let Some(initials) = stamp.get("initials") {
            info.insert("initials".into(), initials.clone());
        }
        if existing.is_some() {
            info.remove("utcDate");
        }
        let mut change = existing
            .cloned()
            .unwrap_or_else(|| json!({"type":"paragraphPropertyChange"}));
        change["info"] = Value::Object(info);
        if existing.is_none() {
            if let Some(previous) = previous.and_then(Value::as_object) {
                let mut properties = previous.clone();
                properties.remove("runProperties");
                properties.remove("runInWithNext");
                change["previousFormatting"] = Value::Object(properties);
            }
        } else if existing.is_some_and(|change| change.get("currentFormatting").is_some()) {
            if let Some(formatting) = result.get("formatting") {
                change["currentFormatting"] = formatting.clone();
            } else {
                change
                    .as_object_mut()
                    .ok_or_else(|| Failure::unsupported(op, "propertyChangeShape"))?
                    .remove("currentFormatting");
            }
        }
        result["propertyChanges"] = json!([change]);
        let inverse = json!({"type":"setParagraphReview","story":op["story"],"blockId":paragraph["paraId"],"expected":review_fields(&result),"review":review_fields(paragraph)});
        return Ok((result, vec![inverse]));
    }
    let inverse = json!({
        "type":"setParagraphProps", "story":op["story"], "blockId":paragraph["paraId"],
        "patch":restore,
        "whenEmpty":if previous.and_then(Value::as_object).is_some_and(Map::is_empty) { "keep" } else { "omit" },
        "expected":op["patch"]
    });
    Ok((result, vec![inverse]))
}

pub(crate) fn contains_identity(value: &Value) -> bool {
    match value {
        Value::Object(fields) => {
            fields.contains_key("propertyChanges")
                || fields.contains_key("info")
                || fields.get("type").and_then(Value::as_str) == Some("inlineSdt")
                || fields.values().any(contains_identity)
        }
        Value::Array(values) => values.iter().any(contains_identity),
        _ => false,
    }
}

pub(crate) fn validate_incoming_inline(value: &Value, op: &Value) -> Result<(), Failure> {
    match value {
        Value::Array(values) => {
            for value in values {
                validate_incoming_inline(value, op)?;
            }
        }
        Value::Object(fields) => {
            let kind = fields.get("type").and_then(Value::as_str);
            let empty = if kind == Some("text") {
                fields
                    .get("text")
                    .and_then(Value::as_str)
                    .is_some_and(str::is_empty)
            } else if matches!(
                kind,
                Some("run" | "insertion" | "deletion" | "moveFrom" | "moveTo")
            ) {
                fields
                    .get("content")
                    .and_then(Value::as_array)
                    .is_some_and(Vec::is_empty)
            } else {
                false
            };
            if empty {
                return Err(Failure::refused(
                    op,
                    "emptyContent",
                    "The content holds an empty run, text node, or revision wrapper.",
                ));
            }
            if kind == Some("text")
                && fields
                    .get("text")
                    .and_then(Value::as_str)
                    .is_some_and(crate::inline::illegal_xml)
            {
                return Err(Failure::refused(
                    op,
                    "invalidText",
                    "The content holds text that cannot be written.",
                ));
            }
            if let Some(content) = fields.get("content") {
                validate_incoming_inline(content, op)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn replace_inline(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    if paragraph.get("content") != op.get("expected") {
        return Err(Failure::refused(
            op,
            "stale",
            format!(
                "{} holds other content than expected.",
                required_string(op, "blockId")?
            ),
        ));
    }
    let content = op
        .get("content")
        .filter(|value| value.is_array())
        .ok_or_else(|| {
            Failure::refused(op, "invalidOperation", "Operation has no content array.")
        })?;
    validate_incoming_inline(content, op)?;
    let mut result = paragraph.clone();
    result["content"] = content.clone();
    if result == *paragraph {
        return Ok((result, vec![]));
    }
    Ok((
        result,
        vec![
            json!({"type":"replaceInline", "story":op["story"], "blockId":paragraph["paraId"],
        "expected":content, "content":paragraph["content"]}),
        ],
    ))
}

fn set_paragraph_review(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    let previous = review_fields(paragraph);
    if op.get("expected") != Some(&previous) {
        return Err(Failure::refused(
            op,
            "stale",
            format!(
                "{} states other review fields than expected.",
                required_string(op, "blockId")?
            ),
        ));
    }
    let review = op.get("review").and_then(Value::as_object).ok_or_else(|| {
        Failure::refused(op, "invalidOperation", "Operation has no review object.")
    })?;
    let mut result = paragraph.clone();
    let fields = result
        .as_object_mut()
        .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?;
    for key in ["formatting", "propertyChanges", "pPrMark"] {
        fields.remove(key);
        if let Some(value) = review.get(key) {
            fields.insert(key.into(), value.clone());
        }
    }
    if result == *paragraph {
        return Ok((result, vec![]));
    }
    Ok((
        result.clone(),
        vec![
            json!({"type":"setParagraphReview", "story":op["story"], "blockId":paragraph["paraId"],
        "expected":review_fields(&result), "review":previous}),
        ],
    ))
}

fn dispatch_one(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let captured = crate::wire::validate_operation(op).map_err(|error| Failure::Refused {
        op_type: error.op_type,
        reason: error.reason.to_string(),
        message: error.message,
    })?;
    if !captured.undefined_fields.is_empty() {
        return Err(Failure::unsupported(op, "ownUndefinedApplication"));
    }
    if let Some(stamp) = op.get("revision") {
        let id = stamp.get("id").and_then(Value::as_i64);
        if !id.is_some_and(|id| (0..=0x7fffffff).contains(&id)) {
            return Err(Failure::refused(
                op,
                "invalidNewId",
                format!(
                    "{} cannot be a revision id.",
                    stamp
                        .get("id")
                        .map(Value::to_string)
                        .unwrap_or_else(|| "undefined".into())
                ),
            ));
        }
        if crate::identity::package_keys(document)
            .contains(&format!("revision:{}", id.unwrap_or(0)))
        {
            return Err(Failure::refused(
                op,
                "idCollision",
                format!(
                    "Revision id {} is already used in the package.",
                    id.unwrap_or(0)
                ),
            ));
        }
    }
    let kind = required_string(op, "type")?;
    if matches!(
        kind,
        "createComment" | "deleteComment" | "restoreCommentState"
    ) {
        return crate::comments::edit(document, op);
    }
    if kind == "insertContent" {
        let incoming = crate::identity::keys(&op["slice"]["content"]);
        let taken = crate::identity::package_keys(document);
        let unique = incoming.iter().collect::<std::collections::BTreeSet<_>>();
        if unique.len() != incoming.len() || incoming.iter().any(|key| taken.contains(key)) {
            return Err(Failure::unsupported(op, "copiedInlineIdentityFreshening"));
        }
    }
    if matches!(kind, "splitBlock" | "joinBlocks") {
        return crate::paragraphs::edit(document, op);
    }
    if matches!(
        kind,
        "insertRow"
            | "deleteRow"
            | "setTableRows"
            | "insertTable"
            | "deleteTable"
            | "setContainerBlocks"
    ) {
        return crate::tables::edit(document, op);
    }
    if matches!(kind, "replaceBlocks" | "insertBlocks" | "deleteBlocks") {
        return crate::blocks::edit(document, op);
    }
    if !matches!(
        kind,
        "setParagraphProps"
            | "replaceInline"
            | "setParagraphReview"
            | "insertText"
            | "deleteRange"
            | "insertContent"
            | "splitInline"
            | "joinInline"
            | "setRunProps"
    ) {
        return Err(Failure::unsupported(op, "operationKind"));
    }
    if matches!(kind, "splitInline" | "joinInline") {
        crate::inline::seam_depth(op)?;
    }
    let address = if matches!(
        kind,
        "insertText" | "insertContent" | "splitInline" | "joinInline"
    ) {
        op.get("at")
    } else if matches!(kind, "deleteRange" | "setRunProps") {
        op.get("from")
    } else {
        None
    };
    if matches!(kind, "deleteRange" | "setRunProps") {
        let from = op.get("from").ok_or_else(|| {
            Failure::refused(op, "invalidOperation", "Operation has no from position.")
        })?;
        let to = op.get("to").ok_or_else(|| {
            Failure::refused(op, "invalidOperation", "Operation has no to position.")
        })?;
        if from.get("story") != to.get("story")
            || !from
                .get("blockId")
                .and_then(Value::as_str)
                .zip(to.get("blockId").and_then(Value::as_str))
                .is_some_and(|(from, to)| from.eq_ignore_ascii_case(to))
        {
            return Err(Failure::refused(
                op,
                "crossBlockRange",
                "A range starts and ends in one paragraph.",
            ));
        }
    }
    let locator = address.map(
        |position| json!({"type":kind,"story":position["story"],"blockId":position["blockId"]}),
    );
    let path = locate(document, locator.as_ref().unwrap_or(op))?;
    let paragraph =
        at_path(document, &path).ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?;
    let (replacement, inverse) = match kind {
        "setParagraphProps" => set_paragraph_props(paragraph, op)?,
        "replaceInline" => replace_inline(paragraph, op)?,
        "setParagraphReview" => set_paragraph_review(paragraph, op)?,
        "insertText" | "deleteRange"
            if op.get("revision").is_some()
                || (kind == "deleteRange"
                    && paragraph["content"].as_array().is_some_and(|items| {
                        items.iter().any(|item| item["type"] == "insertion")
                    })) =>
        {
            crate::tracked::edit(paragraph, op)?
        }
        "insertContent"
            if op["slice"]["content"]
                .as_array()
                .is_some_and(|items| items.iter().any(|item| item["type"] == "insertion")) =>
        {
            crate::tracked::edit(paragraph, op)?
        }
        "insertText" | "deleteRange" | "insertContent" | "splitInline" | "joinInline" => {
            crate::inline::edit(paragraph, op)?
        }
        "setRunProps" => crate::run_props::edit(paragraph, op)?,
        _ => return Err(Failure::unsupported(op, "operationKind")),
    };
    if kind == "replaceInline" {
        let mut outside = document.clone();
        at_path_mut(&mut outside, &path)
            .and_then(Value::as_object_mut)
            .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?
            .remove("content");
        let taken = crate::identity::package_keys(&outside);
        let incoming = crate::identity::keys(&replacement["content"]);
        let unique = incoming.iter().collect::<std::collections::BTreeSet<_>>();
        if unique.len() != incoming.len() || incoming.iter().any(|key| taken.contains(key)) {
            return Err(Failure::refused(
                op,
                "idCollision",
                "The content carries a revision or content-control id already used in the package.",
            ));
        }
    }
    if kind == "setParagraphReview" {
        let review = op.get("review").ok_or_else(|| {
            Failure::refused(op, "invalidOperation", "Operation has no review object.")
        })?;
        if let Some(mark) = review.get("pPrMark") {
            let is_new = paragraph.get("pPrMark") != Some(mark);
            if is_new {
                let mut parent_path = path.clone();
                let index = match parent_path.pop() {
                    Some(Step::Index(index)) => index,
                    _ => return Err(Failure::unsupported(op, "blockListShape")),
                };
                let list = at_path(document, &parent_path)
                    .and_then(Value::as_array)
                    .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
                let cell = parent_path
                    .iter()
                    .any(|step| matches!(step,Step::Field(key) if key=="cells"));
                let permitted = cell
                    && matches!(
                        mark.get("kind").and_then(Value::as_str),
                        Some("ins" | "del")
                    );
                if index + 1 == list.len() && !permitted {
                    return Err(Failure::refused(
                        op,
                        "containerFinalMark",
                        format!(
                            "{} ends its container: there is no next paragraph its mark could join.",
                            op["blockId"].as_str().unwrap_or("")
                        ),
                    ));
                }
            }
        }
        let mut outside = document.clone();
        let target = at_path_mut(&mut outside, &path)
            .and_then(Value::as_object_mut)
            .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?;
        for key in ["formatting", "propertyChanges", "pPrMark"] {
            target.remove(key);
        }
        // Derived sections must mirror the review-free target before identity census.
        let taken = crate::identity::package_keys(&outside);
        let incoming = crate::identity::keys(review);
        let unique = incoming.iter().collect::<std::collections::BTreeSet<_>>();
        if incoming.iter().any(|key| taken.contains(key)) || unique.len() != incoming.len() {
            return Err(Failure::refused(
                op,
                "idCollision",
                "The review fields carry a revision id already used in the package.",
            ));
        }
    }
    let mut modified = Vec::new();
    if replacement != *paragraph {
        modified.push(required_string(paragraph, "paraId")?.to_owned());
    }
    let mut result = document.clone();
    *at_path_mut(&mut result, &path).ok_or_else(|| Failure::unsupported(op, "paragraphShape"))? =
        replacement;
    let revisions = if let Some(stamp) = op.get("revision") {
        let known = crate::identity::keys(paragraph);
        let incoming = crate::identity::keys(
            at_path(&result, &path).ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?,
        );
        incoming
            .iter()
            .filter(|key| !known.contains(key))
            .filter_map(|key| key.strip_prefix("revision:").and_then(|id| id.parse().ok()))
            .filter(|id| stamp.get("id").and_then(Value::as_i64) == Some(*id))
            .collect()
    } else {
        vec![]
    };
    Ok(Applied {
        document: result,
        inverse,
        touched: Touched {
            modified,
            ..Touched::default()
        },
        revisions,
    })
}

/// Atomic batch: no intermediate document escapes after a refusal or unsupported dimension.
pub fn apply_operations(document: &Value, operations: &[Value]) -> Result<Applied, Failure> {
    let mut current = document.clone();
    let mut inverse = Vec::new();
    let mut touched: BTreeMap<String, TouchState> = BTreeMap::new();
    let mut touched_order = Vec::new();
    let mut revisions = Vec::new();
    for op in operations {
        let applied = apply_one(&current, op)?;
        current = applied.document;
        let mut next_inverse = applied.inverse;
        next_inverse.extend(inverse);
        inverse = next_inverse;
        for (state, ids) in [
            (TouchState::Modified, applied.touched.modified),
            (TouchState::Inserted, applied.touched.inserted),
            (TouchState::Removed, applied.touched.removed),
        ] {
            for id in ids {
                let previous = touched.get(&id).copied();
                let next = combine_touch(previous, state);
                match next {
                    Some(next) => {
                        if previous.is_none() {
                            touched_order.push(id.clone());
                        }
                        touched.insert(id, next);
                    }
                    None => {
                        touched.remove(&id);
                        touched_order.retain(|previous| previous != &id);
                    }
                }
            }
        }
        revisions.extend(applied.revisions);
    }
    let ids = |state| {
        touched_order
            .iter()
            .filter(|id| touched.get(*id) == Some(&state))
            .cloned()
            .collect()
    };
    Ok(Applied {
        document: current,
        inverse,
        touched: Touched {
            modified: ids(TouchState::Modified),
            inserted: ids(TouchState::Inserted),
            removed: ids(TouchState::Removed),
        },
        revisions,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TouchState {
    Modified,
    Inserted,
    Removed,
}

fn combine_touch(previous: Option<TouchState>, next: TouchState) -> Option<TouchState> {
    use TouchState::{Inserted, Modified, Removed};
    match (previous, next) {
        (None, state) => Some(state),
        (Some(Modified), Modified | Inserted) => Some(Modified),
        (Some(Modified), Removed) => Some(Removed),
        (Some(Inserted), Modified | Inserted) => Some(Inserted),
        (Some(Inserted), Removed) => None,
        (Some(Removed), Modified | Inserted) => Some(Modified),
        (Some(Removed), Removed) => Some(Removed),
    }
}

/// Mirror TS withBodyContent: preserve section metadata and derive only content.
pub(crate) fn refresh_sections(document: &mut Value, op: &Value) -> Result<(), Failure> {
    let body = document
        .pointer_mut("/package/document")
        .ok_or_else(|| Failure::unsupported(op, "documentShape"))?;
    let Some(previous) = body.get("sections").and_then(Value::as_array) else {
        return Ok(());
    };
    let content = body
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "documentShape"))?;
    let mut groups = Vec::new();
    let mut current = Vec::new();
    for block in content {
        current.push(block.clone());
        if block.get("type").and_then(Value::as_str) == Some("paragraph")
            && block.get("sectionProperties").is_some()
        {
            groups.push(std::mem::take(&mut current));
        }
    }
    if !current.is_empty() || groups.is_empty() {
        groups.push(current);
    }
    if groups.len() != previous.len() {
        return Err(Failure::unsupported(op, "sectionCountChange"));
    }
    let mut sections = Vec::new();
    for (section, content) in previous.iter().zip(groups) {
        let mut next = section.clone();
        next["content"] = Value::Array(content);
        sections.push(next);
    }
    body["sections"] = Value::Array(sections);
    Ok(())
}

fn apply_one(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let mut applied = dispatch_one(document, op)?;
    refresh_sections(&mut applied.document, op)?;
    Ok(applied)
}
