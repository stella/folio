//! Paragraph block-list edits, with the same anchored replacement inverses as TS.
use crate::apply::{
    Applied, Path, Step, Touched, at_path, at_path_mut, contains_identity, locate, required_string,
};
use crate::refusal::Failure;
use serde_json::{Value, json};
use std::collections::{BTreeMap, BTreeSet};

fn array<'a>(op: &'a Value, key: &str) -> Result<&'a Vec<Value>, Failure> {
    op.get(key).and_then(Value::as_array).ok_or_else(|| {
        Failure::refused(
            op,
            "invalidOperation",
            format!("Operation has no {key} array."),
        )
    })
}

fn index_and_list(path: &[Step], op: &Value) -> Result<(usize, Path), Failure> {
    let Some(Step::Index(index)) = path.last() else {
        return Err(Failure::unsupported(op, "blockListShape"));
    };
    Ok((*index, path[..path.len() - 1].to_vec()))
}

fn locator(op: &Value, id: &str) -> Value {
    json!({"type":op["type"],"story":op["story"],"blockId":id})
}

fn ids(values: &[Value]) -> BTreeSet<String> {
    values
        .iter()
        .filter_map(|value| value.get("paraId").and_then(Value::as_str))
        .map(str::to_ascii_uppercase)
        .collect()
}

fn is_para_id(id: &str) -> bool {
    id.len() == 8
        && id.bytes().all(|c| c.is_ascii_hexdigit())
        && u32::from_str_radix(id, 16).is_ok_and(|id| id > 0 && id < 0x80000000)
}

fn touched(before: &[Value], after: &[Value]) -> Touched {
    let old: BTreeMap<_, _> = before
        .iter()
        .filter_map(|p| {
            p.get("paraId")
                .and_then(Value::as_str)
                .map(|id| (id.to_ascii_uppercase(), p))
        })
        .collect();
    let new_ids = ids(after);
    let mut result = Touched::default();
    for p in after {
        let Some(id) = p.get("paraId").and_then(Value::as_str) else {
            continue;
        };
        match old.get(&id.to_ascii_uppercase()) {
            None => result.inserted.push(id.into()),
            Some(previous) if *previous != p => result.modified.push(id.into()),
            _ => {}
        }
    }
    for p in before {
        if let Some(id) = p.get("paraId").and_then(Value::as_str) {
            if !new_ids.contains(&id.to_ascii_uppercase()) {
                result.removed.push(id.into());
            }
        }
    }
    result
}

fn empty_record(value: &Value) -> bool {
    match value {
        Value::Object(fields) => {
            let kind = fields.get("type").and_then(Value::as_str);
            (kind == Some("text") && fields.get("text").and_then(Value::as_str) == Some(""))
                || (matches!(
                    kind,
                    Some("run" | "insertion" | "deletion" | "moveFrom" | "moveTo")
                ) && fields
                    .get("content")
                    .and_then(Value::as_array)
                    .is_some_and(Vec::is_empty))
                || fields.get("content").is_some_and(empty_record)
        }
        Value::Array(values) => values.iter().any(empty_record),
        _ => false,
    }
}

fn replace(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let expected = array(op, "expected")?;
    let blocks = array(op, "blocks")?;
    if expected.is_empty() || blocks.is_empty() {
        return Err(Failure::refused(
            op,
            "emptyBlockList",
            "A replacement names the paragraphs it replaces and the ones it puts in their place.",
        ));
    }
    let mut found = Vec::new();
    for paragraph in expected {
        let Some(id) = paragraph.get("paraId").and_then(Value::as_str) else {
            return Err(Failure::refused(
                op,
                "blockNotFound",
                "A replaced paragraph has no id.",
            ));
        };
        found.push(locate(document, &locator(op, id))?);
    }
    let (start, list_path) = index_and_list(&found[0], op)?;
    for (offset, path) in found.iter().enumerate() {
        let (index, list) = index_and_list(path, op)?;
        if list != list_path || index != start + offset {
            return Err(Failure::refused(
                op,
                "notAdjacent",
                "The replaced paragraphs are not adjacent.",
            ));
        }
    }
    let before = found
        .iter()
        .map(|path| {
            at_path(document, path)
                .cloned()
                .ok_or_else(|| Failure::unsupported(op, "blockListShape"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    if before != *expected {
        return Err(Failure::refused(
            op,
            "stale",
            "The paragraphs to replace have changed.",
        ));
    }
    if blocks
        .iter()
        .any(|p| p.get("paraId").and_then(Value::as_str).is_none())
    {
        return Err(Failure::refused(
            op,
            "invalidBlockId",
            "A replacement paragraph has no id.",
        ));
    }
    if blocks.iter().any(empty_record) {
        return Err(Failure::refused(
            op,
            "emptyContent",
            "A replacement paragraph holds an empty run, text node, or revision wrapper.",
        ));
    }
    let old_ids = ids(&before);
    let new_ids = ids(blocks);
    let created = blocks
        .iter()
        .filter(|p| !old_ids.contains(&p["paraId"].as_str().unwrap_or("").to_ascii_uppercase()))
        .chain(before.iter().filter(|p| {
            !new_ids.contains(&p["paraId"].as_str().unwrap_or("").to_ascii_uppercase())
        }));
    if created
        .into_iter()
        .any(|p| !is_para_id(p["paraId"].as_str().unwrap_or("")))
    {
        return Err(Failure::refused(
            op,
            "invalidBlockId",
            "A paragraph the replacement adds or removes has no usable id.",
        ));
    }
    if op.get("sectionView").is_some() {
        return Err(Failure::unsupported(op, "sectionViewRestoration"));
    }
    if op.get("sectionBoundaries").and_then(Value::as_str) != Some("replace") {
        let inner_break = |ps: &[Value]| {
            ps.iter()
                .take(ps.len().saturating_sub(1))
                .any(|p| p.get("sectionProperties").is_some())
        };
        if inner_break(&before)
            || inner_break(blocks)
            || before.last().and_then(|p| p.get("sectionProperties"))
                != blocks.last().and_then(|p| p.get("sectionProperties"))
        {
            return Err(Failure::refused(
                op,
                "sectionBoundary",
                "A replacement keeps every section break where it is.",
            ));
        }
    }
    if blocks.iter().any(contains_identity) {
        return Err(Failure::unsupported(op, "blockIdentityValidation"));
    }
    // Package has document and secondary stories, not just container edges.
    let mut global_ids = Vec::new();
    let mut package = document["package"].clone();
    if let Some(body) = package.get_mut("document").and_then(Value::as_object_mut) {
        body.remove("sections");
    }
    collect_paragraph_ids(&package, &mut global_ids);
    let mut remaining: BTreeMap<String, usize> = BTreeMap::new();
    for id in global_ids {
        *remaining.entry(id).or_default() += 1;
    }
    for p in &before {
        if let Some(count) =
            remaining.get_mut(&p["paraId"].as_str().unwrap_or("").to_ascii_uppercase())
        {
            *count = count.saturating_sub(1);
        }
    }
    let mut incoming = BTreeSet::new();
    for p in blocks {
        let id = p["paraId"].as_str().unwrap_or("").to_ascii_uppercase();
        if remaining.get(&id).copied().unwrap_or(0) > 0 || !incoming.insert(id) {
            return Err(Failure::refused(
                op,
                "idCollision",
                "A replacement paragraph id is already used in the package.",
            ));
        }
    }
    let mut result = document.clone();
    let list = at_path_mut(&mut result, &list_path)
        .and_then(Value::as_array_mut)
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    list.splice(start..start + before.len(), blocks.iter().cloned());
    let mut inverse =
        json!({"type":"replaceBlocks","story":op["story"],"expected":blocks,"blocks":before});
    if op.get("sectionBoundaries").and_then(Value::as_str) == Some("replace") {
        inverse["sectionBoundaries"] = json!("replace");
    }
    Ok(Applied {
        document: result,
        inverse: vec![inverse],
        touched: touched(&before, blocks),
        revisions: vec![],
    })
}

fn collect_paragraph_ids(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::Object(fields) => {
            if fields.get("type").and_then(Value::as_str) == Some("paragraph") {
                if let Some(id) = fields.get("paraId").and_then(Value::as_str) {
                    out.push(id.to_ascii_uppercase());
                }
                return;
            }
            for child in fields.values() {
                collect_paragraph_ids(child, out);
            }
        }
        Value::Array(values) => {
            for value in values {
                collect_paragraph_ids(value, out);
            }
        }
        _ => {}
    }
}

fn insert(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let blocks = array(op, "blocks")?;
    if blocks.is_empty() {
        return Err(Failure::refused(
            op,
            "emptyBlockList",
            "No paragraphs to insert.",
        ));
    }
    let at = op.get("at").ok_or_else(|| {
        Failure::refused(op, "invalidOperation", "Operation has no insertion point.")
    })?;
    let side = at.get("type").and_then(Value::as_str);
    if !matches!(side, Some("before" | "after")) {
        return Err(Failure::refused(
            op,
            "structureMismatch",
            "Insert before or after a paragraph.",
        ));
    }
    let id = required_string(at, "blockId")?;
    let path = locate(document, &locator(op, id)).map_err(|_| {
        Failure::refused(op, "blockNotFound", "The insertion anchor does not exist.")
    })?;
    if blocks.iter().any(|p| p.get("sectionProperties").is_some()) {
        return Err(Failure::refused(
            op,
            "untrackable",
            "Inserting section breaks requires section operations.",
        ));
    }
    let (index, list_path) = index_and_list(&path, op)?;
    let list = at_path(document, &list_path)
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    let current = list
        .get(index)
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    let following = list
        .get(index + 1)
        .filter(|p| p.get("type").and_then(Value::as_str) == Some("paragraph"));
    let before = side == Some("before") || following.is_some();
    let anchor = if side == Some("after") {
        following.unwrap_or(current)
    } else {
        current
    };
    if op.get("revision").is_some() {
        if !before {
            return Err(Failure::refused(
                op,
                "untrackable",
                "Tracked insertion requires a following paragraph in the same block list.",
            ));
        }
        return Err(Failure::unsupported(op, "trackedBlockInsertion"));
    }
    let mut replacement = blocks.clone();
    if before {
        replacement.push(anchor.clone());
    } else {
        replacement.insert(0, anchor.clone());
    }
    let mut applied = replace(
        document,
        &json!({"type":"replaceBlocks","story":op["story"],"expected":[anchor],"blocks":replacement}),
    )?;
    applied.touched = Touched {
        inserted: blocks
            .iter()
            .filter_map(|p| p.get("paraId").and_then(Value::as_str).map(str::to_owned))
            .collect(),
        ..Touched::default()
    };
    Ok(applied)
}

pub(crate) fn edit(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let result = match op.get("type").and_then(Value::as_str) {
        Some("replaceBlocks") => replace(document, op),
        Some("insertBlocks") => insert(document, op),
        Some("deleteBlocks") => delete(document, op),
        _ => Err(Failure::unsupported(op, "operationKind")),
    };
    result.map_err(|failure| failure.with_op(op))
}

fn plain_length(paragraph: &Value, op: &Value) -> Result<usize, Failure> {
    let mut length = 0;
    let content = paragraph
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?;
    for run in content {
        if run.get("type").and_then(Value::as_str) != Some("run") {
            return Err(Failure::unsupported(op, "terminalDeletionInlineShape"));
        }
        for leaf in run
            .get("content")
            .and_then(Value::as_array)
            .ok_or_else(|| Failure::unsupported(op, "runContentShape"))?
        {
            if leaf.get("type").and_then(Value::as_str) != Some("text") {
                return Err(Failure::unsupported(op, "terminalDeletionInlineShape"));
            }
            length += leaf
                .get("text")
                .and_then(Value::as_str)
                .ok_or_else(|| Failure::unsupported(op, "textLeafShape"))?
                .encode_utf16()
                .count();
        }
    }
    Ok(length)
}

fn delete(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let selected_ids = array(op, "blockIds")?;
    if selected_ids.is_empty() {
        return Err(Failure::refused(
            op,
            "emptyBlockList",
            "No paragraphs to delete.",
        ));
    }
    let mut found = Vec::new();
    for id in selected_ids {
        let id = id.as_str().ok_or_else(|| {
            Failure::refused(op, "invalidOperation", "A block id must be a string.")
        })?;
        found.push(locate(document, &locator(op, id)).map_err(|_| {
            Failure::refused(op, "blockNotFound", "A selected paragraph does not exist.")
        })?);
    }
    let (start, list_path) = index_and_list(&found[0], op)?;
    for (offset, path) in found.iter().enumerate() {
        let (index, list) = index_and_list(path, op)?;
        if list != list_path || index != start + offset {
            return Err(Failure::refused(
                op,
                "notAdjacent",
                "Selected paragraphs must be adjacent in one block list.",
            ));
        }
    }
    let list = at_path(document, &list_path)
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    let end = start + found.len();
    let selected = list
        .get(start..end)
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    let following = list.get(end);
    let terminal = end == list.len();
    if !terminal
        && following
            .is_some_and(|block| block.get("type").and_then(Value::as_str) != Some("paragraph"))
    {
        return Err(Failure::refused(
            op,
            "untrackable",
            "Deletion requires a following paragraph or a container-final paragraph.",
        ));
    }
    let previous = if terminal {
        start
            .checked_sub(1)
            .and_then(|index| list.get(index))
            .filter(|p| p.get("type").and_then(Value::as_str) == Some("paragraph"))
    } else {
        None
    };
    let mut affected = Vec::new();
    if let Some(previous) = previous {
        affected.push(previous.clone());
    }
    affected.extend_from_slice(selected);
    if affected
        .iter()
        .any(|p| p.get("sectionProperties").is_some())
    {
        return Err(Failure::refused(
            op,
            "untrackable",
            "Deleting section boundaries requires section operations.",
        ));
    }
    if op.get("revision").is_some() {
        return Err(Failure::unsupported(op, "trackedBlockDeletion"));
    }
    let last = selected
        .last()
        .ok_or_else(|| Failure::unsupported(op, "blockListShape"))?;
    let mut expected = if following.is_some() {
        selected.to_vec()
    } else {
        affected.clone()
    };
    let blocks = if let Some(following) = following {
        expected.push(following.clone());
        vec![following.clone()]
    } else {
        let mut survivor = last.clone();
        survivor["content"] = previous
            .map(|p| p["content"].clone())
            .unwrap_or_else(|| json!([]));
        if let Some(previous) = previous {
            if plain_length(previous, op)? > 0 {
                let mut formatting = previous
                    .get("formatting")
                    .and_then(Value::as_object)
                    .cloned()
                    .unwrap_or_default();
                for key in ["runProperties", "runInWithNext"] {
                    formatting.remove(key);
                    if let Some(value) = last.get("formatting").and_then(|f| f.get(key)) {
                        formatting.insert(key.into(), value.clone());
                    }
                }
                if formatting.is_empty() {
                    survivor
                        .as_object_mut()
                        .ok_or_else(|| Failure::unsupported(op, "paragraphShape"))?
                        .remove("formatting");
                } else {
                    survivor["formatting"] = Value::Object(formatting);
                }
            }
        }
        vec![survivor]
    };
    let mut applied = replace(
        document,
        &json!({"type":"replaceBlocks","story":op["story"],"expected":expected,"blocks":blocks}),
    )?;
    applied.touched = Touched {
        modified: if terminal {
            vec![last["paraId"].as_str().unwrap_or("").into()]
        } else {
            vec![]
        },
        inserted: vec![],
        removed: affected
            .iter()
            .filter(|p| !terminal || *p != last)
            .filter_map(|p| p.get("paraId").and_then(Value::as_str).map(str::to_owned))
            .collect(),
    };
    Ok(applied)
}
