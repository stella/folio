//! Schema-10 comment semantics from `origin/feat/canonical-comments`.
//! Source: 617c0f4703a6cf43adb7ecda34366c29b93cc60f, ops/comments.ts.
//! This module is isolated spike code. JS Map transport is deliberately not
//! invented: the JSON entry point refuses it; the semantic entry point carries
//! ordered relationships as a typed sidecar. Main-story plain runs and top-level
//! comment anchors are supported; tracked/nested inline scaffolding is not.

use crate::apply::{Applied, Path, Touched, at_path, at_path_mut, locate, paragraph_paths};
use crate::refusal::Failure;
use serde_json::{Map, Value, json};
use std::collections::{BTreeMap, BTreeSet};

pub const SOURCE_COMMIT: &str = "617c0f4703a6cf43adb7ecda34366c29b93cc60f";
const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const EXTENDED_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const MAX_ID: u64 = 2_147_483_647;

/// A JS Map is not JSON. Preserve its ordered entries and own presence explicitly.
#[derive(Clone, Debug, PartialEq)]
pub enum RelationshipState {
    Absent,
    Undefined,
    Present(Vec<(String, Value)>),
}

impl RelationshipState {
    fn entries(&self) -> &[(String, Value)] {
        match self {
            Self::Present(entries) => entries,
            _ => &[],
        }
    }
    fn presence(&self) -> &str {
        match self {
            Self::Absent => "absent",
            Self::Undefined => "undefined",
            Self::Present(_) => "present",
        }
    }
}

pub struct CommentApplied {
    pub applied: Applied,
    pub relationships: RelationshipState,
}

fn fail(op: &Value, message: &str) -> Failure {
    Failure::refused(op, "structureMismatch", message)
}

fn valid_id(value: &Value) -> Option<u64> {
    value.as_u64().filter(|id| *id <= MAX_ID)
}
fn kind(node: &Value) -> &str {
    node["type"].as_str().unwrap_or("")
}
fn anchor_id(node: &Value) -> Option<u64> {
    matches!(
        kind(node),
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    )
    .then(|| valid_id(&node["id"]))
    .flatten()
}
fn is_anchor(node: &Value) -> bool {
    matches!(
        kind(node),
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    )
}
fn width(node: &Value) -> usize {
    if kind(node) == "run" {
        crate::inline::list_width(std::slice::from_ref(node))
    } else if kind(node) == "commentReference" {
        1
    } else {
        0
    }
}
fn content(paragraph: &Value) -> &[Value] {
    paragraph["content"]
        .as_array()
        .expect("validated paragraph content")
}
fn main_paths(document: &Value) -> Vec<Path> {
    let mut paths = Vec::new();
    let mut base = vec![
        crate::apply::Step::Field("package".into()),
        crate::apply::Step::Field("document".into()),
    ];
    paragraph_paths(&document["package"]["document"], &mut base, &mut paths);
    paths
}
fn validate_content(nodes: &[Value], op: &Value) -> Result<(), Failure> {
    for node in nodes {
        if kind(node) == "run" {
            crate::inline::validate_plain(std::slice::from_ref(node), op, false)?;
            if !crate::identity::keys(node).is_empty() {
                return Err(Failure::unsupported(op, "identifiedCommentScaffolding"));
            }
        } else if !is_anchor(node) {
            return Err(Failure::unsupported(
                op,
                "commentInlineContainersAndSpecialLeaves",
            ));
        } else if anchor_id(node).is_none() {
            return Err(fail(op, "A comment anchor has no definition."));
        }
    }
    Ok(())
}
fn comments(document: &Value, op: &Value) -> Result<Vec<Value>, Failure> {
    match document.pointer("/package/document/comments") {
        None => Ok(Vec::new()),
        Some(Value::Array(records)) => Ok(records.clone()),
        _ => Err(Failure::unsupported(op, "commentListPresence")),
    }
}

#[derive(Clone)]
struct Anchor {
    block: String,
    node: Value,
    after_offset: usize,
    after_zero: usize,
}

fn anchors(document: &Value, op: &Value) -> Result<Vec<Anchor>, Failure> {
    let mut out = Vec::new();
    for path in main_paths(document) {
        let paragraph = at_path(document, &path).expect("paragraph path exists");
        let nodes = paragraph
            .get("content")
            .and_then(Value::as_array)
            .ok_or_else(|| Failure::unsupported(op, "paragraphContentShape"))?;
        validate_content(nodes, op)?;
        let mut offset = 0;
        let mut zero = 0;
        for node in nodes {
            let length = width(node);
            if length == 0 {
                zero += 1;
            } else {
                offset += length;
                zero = 0;
            }
            if is_anchor(node) {
                out.push(Anchor {
                    block: paragraph["paraId"].as_str().unwrap_or("").into(),
                    node: node.clone(),
                    after_offset: offset,
                    after_zero: zero,
                });
            }
        }
    }
    // Other story anchors cannot be silently excluded from the activation contract.
    fn holds_anchor(value: &Value) -> bool {
        if is_anchor(value) {
            return true;
        }
        match value {
            Value::Array(items) => items.iter().any(holds_anchor),
            Value::Object(fields) => fields.values().any(holds_anchor),
            _ => false,
        }
    }
    if document["package"].as_object().is_some_and(|fields| {
        fields
            .iter()
            .any(|(key, value)| key != "document" && holds_anchor(value))
    }) {
        return Err(Failure::unsupported(op, "secondaryCommentAnchors"));
    }
    Ok(out)
}

fn issue(document: &Value, op: &Value) -> Result<(), Failure> {
    let records = comments(document, op)?;
    let mut by_id = BTreeMap::new();
    for record in &records {
        let id = valid_id(&record["id"])
            .ok_or_else(|| fail(op, "Comment identities must be unique bounded integers."))?;
        if by_id.insert(id, record).is_some() {
            return Err(fail(
                op,
                "Comment identities must be unique bounded integers.",
            ));
        }
    }
    let spans = anchors(document, op)?;
    if spans
        .iter()
        .any(|a| !by_id.contains_key(&anchor_id(&a.node).expect("validated anchor")))
    {
        return Err(fail(op, "A comment anchor has no definition."));
    }
    for (id, record) in &by_id {
        let own: Vec<_> = spans
            .iter()
            .enumerate()
            .filter(|(_, a)| anchor_id(&a.node) == Some(*id))
            .collect();
        let starts: Vec<_> = own
            .iter()
            .filter(|(_, a)| kind(&a.node) == "commentRangeStart")
            .collect();
        let ends: Vec<_> = own
            .iter()
            .filter(|(_, a)| kind(&a.node) == "commentRangeEnd")
            .collect();
        let references = own
            .iter()
            .filter(|(_, a)| kind(&a.node) == "commentReference")
            .count();
        if starts.len() != ends.len() || starts.len() > 1 || references > 1 {
            return Err(fail(op, "Comment anchors are unbalanced or duplicated."));
        }
        if starts
            .first()
            .zip(ends.first())
            .is_some_and(|(start, end)| start.0 > end.0)
        {
            return Err(fail(
                op,
                "Comment range boundaries must be ordered in one story.",
            ));
        }
        let parent = record.get("parentId").and_then(valid_id);
        if record.get("parentId").is_some() && parent.is_none() {
            return Err(fail(op, "A comment parent identity is invalid."));
        }
        if !parent.is_some_and(|parent| by_id.contains_key(&parent)) && own.is_empty() {
            return Err(fail(
                op,
                "A root or revision-associated comment needs an owned source anchor.",
            ));
        }
        let mut visited = BTreeSet::from([*id]);
        let mut current = parent.and_then(|parent| by_id.get(&parent));
        while let Some(parent_record) = current {
            let parent_id = valid_id(&parent_record["id"]).expect("validated id");
            if !visited.insert(parent_id) {
                return Err(fail(op, "Comment parent relations cannot cycle."));
            }
            current = parent_record
                .get("parentId")
                .and_then(valid_id)
                .and_then(|parent| by_id.get(&parent));
        }
    }
    Ok(())
}

fn locate_position(document: &Value, position: &Value, op: &Value) -> Result<Path, Failure> {
    locate(
        document,
        &json!({"type":op["type"],"story":position["story"],"blockId":position["blockId"]}),
    )
    .map_err(|e| e.with_op(op))
}

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Gap {
    offset: usize,
    zero: usize,
}

fn gap(nodes: &[Value], position: &Value, op: &Value) -> Result<Gap, Failure> {
    let offset = position["offset"]
        .as_u64()
        .and_then(|n| usize::try_from(n).ok())
        .filter(|n| *n <= nodes.iter().map(width).sum())
        .ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOffset",
                "The offset is not a position in the paragraph.",
            )
        })?;
    let mut current = 0;
    let mut zeros = Vec::new();
    for node in nodes {
        let length = width(node);
        if length == 0 && current == offset {
            zeros.push(kind(node));
        }
        if kind(node) == "run" && current <= offset && offset <= current + length {
            crate::inline::coordinate(
                std::slice::from_ref(node),
                &json!({"offset":offset-current}),
                op,
            )?;
        }
        current += length;
    }
    let zero = match position.get("zeroWidthBefore") {
        None => zeros
            .iter()
            .position(|kind| *kind == "commentRangeStart")
            .unwrap_or(zeros.len()),
        Some(value) => value
            .as_u64()
            .and_then(|n| usize::try_from(n).ok())
            .filter(|n| *n <= zeros.len())
            .ok_or_else(|| {
                Failure::refused(
                    op,
                    "invalidOffset",
                    "The zero-width coordinate is unavailable.",
                )
            })?,
    };
    Ok(Gap { offset, zero })
}

fn insert(nodes: &[Value], at: Gap, inserted: &[Value]) -> Vec<Value> {
    let mut left = Vec::new();
    let mut right = Vec::new();
    let mut offset = 0;
    let mut zero = 0;
    for node in nodes {
        let length = width(node);
        if length == 0 {
            if (Gap { offset, zero }) < at {
                left.push(node.clone());
            } else {
                right.push(node.clone());
            }
            zero += 1;
            continue;
        }
        if kind(node) == "run" && offset < at.offset && at.offset < offset + length {
            let pieces =
                crate::inline::partition(std::slice::from_ref(node), &[at.offset - offset]);
            left.extend(pieces[0].clone());
            right.extend(pieces[1].clone());
        } else if offset < at.offset {
            left.push(node.clone());
        } else {
            right.push(node.clone());
        }
        offset += length;
        zero = 0;
    }
    left.extend_from_slice(inserted);
    left.extend(right);
    left
}

fn run_fields(node: &Value) -> Map<String, Value> {
    node.as_object()
        .expect("run object")
        .iter()
        .filter(|(key, _)| key.as_str() != "content")
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect()
}
fn merge_alike(left: &[Value], right: &[Value]) -> Vec<Value> {
    let depth = match left.last().zip(right.first()) {
        Some((a, b)) if kind(a) == "run" && kind(b) == "run" && run_fields(a) == run_fields(b) => 2,
        _ => 0,
    };
    crate::inline::merge(left, right, depth, false).expect("alike plain runs merge")
}
fn owned_seams(nodes: &[Value], ids: &BTreeSet<u64>) -> BTreeSet<usize> {
    let mut references = 0;
    let mut offset = 0;
    let mut seams = BTreeSet::new();
    for node in nodes {
        if anchor_id(node).is_some_and(|id| ids.contains(&id)) {
            seams.insert(offset - references);
            if kind(node) == "commentReference" {
                references += 1;
            }
        }
        offset += width(node);
    }
    seams
}
fn normalized(nodes: &[Value], ids: &BTreeSet<u64>, seams: &BTreeSet<usize>) -> Vec<Value> {
    let mut result = Vec::new();
    let mut offset = 0;
    for node in nodes {
        if anchor_id(node).is_some_and(|id| ids.contains(&id)) {
            continue;
        }
        if seams.contains(&offset) {
            result = merge_alike(&result, std::slice::from_ref(node));
        } else {
            result.push(node.clone());
        }
        offset += width(node);
    }
    result
}

fn is_relation(value: &Value) -> bool {
    matches!(value["type"].as_str(), Some(COMMENTS_REL | EXTENDED_REL))
}
fn capture(
    document: &Value,
    relations: &RelationshipState,
    ids: &BTreeSet<u64>,
    addressed: &[Value],
) -> Value {
    let records = document
        .pointer("/package/document/comments")
        .and_then(Value::as_array);
    let anchors: Vec<_> = addressed.iter().filter_map(|entry| {
        let path = locate_position(document, entry, &json!({"type":"restoreCommentState"})).ok()?;
        Some(json!({"story":entry["story"],"blockId":entry["blockId"],"content":at_path(document,&path)?["content"]}))
    }).collect();
    json!({
        "relationshipPresence":relations.presence(),
        "relationships":relations.entries().iter().enumerate().filter(|(_,(_,r))|is_relation(r))
            .map(|(index,(key,relationship))|json!({"index":index,"key":key,"relationship":relationship})).collect::<Vec<_>>(),
        "listPresence":if document.pointer("/package/document/comments").is_some(){"present"}else{"absent"},
        "records":records.into_iter().flatten().enumerate().filter(|(_,record)|valid_id(&record["id"]).is_some_and(|id|ids.contains(&id)))
            .map(|(index,comment)|json!({"index":index,"comment":comment})).collect::<Vec<_>>(),
        "anchors":anchors,
    })
}

fn reconcile(document: &Value, previous: &RelationshipState) -> RelationshipState {
    let comments = document
        .pointer("/package/document/comments")
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .unwrap_or(&[]);
    let mut relationships = previous.entries().to_vec();
    let mut next = relationships
        .iter()
        .filter_map(|(key, _)| {
            key.strip_prefix("rId")
                .filter(|digits| !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()))
                .and_then(|digits| digits.parse::<u64>().ok())
        })
        .max()
        .unwrap_or(0)
        + 1;
    let extended = comments
        .iter()
        .any(|record| record.get("parentId").is_some() || record.get("done").is_some());
    for (relation, target, needed) in [
        (COMMENTS_REL, "comments.xml", !comments.is_empty()),
        (EXTENDED_REL, "commentsExtended.xml", extended),
    ] {
        if needed
            && !relationships
                .iter()
                .any(|(_, value)| value["type"] == relation)
        {
            let id = format!("rId{next}");
            next += 1;
            relationships.push((id.clone(), json!({"id":id,"type":relation,"target":target})));
        }
        if relation == EXTENDED_REL && !needed {
            relationships.retain(|(_, value)| value["type"] != relation);
        }
    }
    if relationships == previous.entries() {
        previous.clone()
    } else {
        RelationshipState::Present(relationships)
    }
}

fn finish(
    before: &Value,
    after: Value,
    old_rel: &RelationshipState,
    new_rel: RelationshipState,
    op: &Value,
    id_order: &[u64],
) -> Result<CommentApplied, Failure> {
    let ids: BTreeSet<_> = id_order.iter().copied().collect();
    issue(&after, op)?;
    // Package paragraph identities include comments but never derived section copies.
    fn census(value: &Value, out: &mut BTreeMap<String, usize>) {
        match value {
            Value::Array(items) => {
                for item in items {
                    census(item, out);
                }
            }
            Value::Object(fields) => {
                if kind(value) == "paragraph" {
                    if let Some(id) = value["paraId"].as_str() {
                        *out.entry(id.to_ascii_uppercase()).or_default() += 1;
                    }
                }
                for child in fields.values() {
                    census(child, out);
                }
            }
            _ => {}
        }
    }
    let mut package = after["package"].clone();
    if let Some(body) = package["document"].as_object_mut() {
        body.remove("sections");
    }
    let mut counts = BTreeMap::new();
    census(&package, &mut counts);
    if let Some((id, count)) = counts.into_iter().find(|(_, count)| *count > 1) {
        return Err(Failure::refused(
            op,
            "duplicateBlockId",
            format!("{count} paragraphs in the package are {id}."),
        ));
    }
    let mut record_counts = BTreeMap::new();
    for key in crate::identity::package_keys(&after) {
        *record_counts.entry(key).or_insert(0) += 1;
    }
    if let Some((key, count)) = record_counts.into_iter().find(|(_, count)| *count > 1) {
        return Err(Failure::refused(
            op,
            "duplicateRecordId",
            format!("{count} records in the package carry {key}."),
        ));
    }
    let mut addressed = Vec::new();
    for path in main_paths(before) {
        let old = at_path(before, &path).expect("existing paragraph");
        if at_path(&after, &path).is_none_or(|next| next["content"] != old["content"]) {
            addressed.push(json!({"story":"main","blockId":old["paraId"]}));
        }
    }
    let prior = capture(before, old_rel, &ids, &addressed);
    let next = capture(&after, &new_rel, &ids, &addressed);
    let inverse = if prior == next {
        vec![]
    } else {
        vec![json!({"type":"restoreCommentState","ids":id_order,
        "scaffoldIds":{"revision":[],"control":[]},"expected":next,"state":prior})]
    };
    let modified = if inverse.is_empty() {
        vec![]
    } else {
        addressed
            .iter()
            .filter_map(|entry| entry["blockId"].as_str().map(str::to_owned))
            .collect()
    };
    Ok(CommentApplied {
        applied: Applied {
            document: if inverse.is_empty() {
                before.clone()
            } else {
                after
            },
            inverse,
            touched: Touched {
                modified,
                ..Touched::default()
            },
            revisions: vec![],
        },
        relationships: new_rel,
    })
}

fn with_comments(
    document: &mut Value,
    records: Vec<Value>,
    presence: &str,
    op: &Value,
) -> Result<(), Failure> {
    let body = document
        .pointer_mut("/package/document")
        .and_then(Value::as_object_mut)
        .ok_or_else(|| Failure::unsupported(op, "documentShape"))?;
    match presence {
        "present" => {
            body.insert("comments".into(), Value::Array(records));
        }
        "absent" => {
            body.remove("comments");
        }
        "undefined" => return Err(Failure::unsupported(op, "commentListOwnUndefined")),
        _ => return Err(fail(op, "The comment inverse payload is invalid.")),
    }
    Ok(())
}

fn create(
    document: &Value,
    op: &Value,
    relations: &RelationshipState,
) -> Result<CommentApplied, Failure> {
    let records = comments(document, op)?;
    let mut comment = op["comment"].clone();
    if !comment["author"].is_string() || !comment["content"].is_array() || !op["anchor"].is_object()
    {
        return Err(fail(op, "The comment creation payload is invalid."));
    }
    let id = valid_id(&comment["id"])
        .filter(|id| {
            !records.iter().any(|record| record["id"] == *id)
                && !crate::identity::package_keys(document).contains(&format!("revision:{id}"))
        })
        .ok_or_else(|| {
            Failure::refused(
                op,
                "idCollision",
                "The fresh comment identity is invalid or already used.",
            )
        })?;
    if comment.get("parentId").is_some() {
        return Err(fail(
            op,
            "Comment parent ownership belongs to the anchor discriminator.",
        ));
    }
    let paragraphs = comment["content"].as_array().expect("checked content");
    if paragraphs.is_empty()
        || paragraphs.iter().any(|paragraph| {
            paragraph["paraId"].as_str().is_none_or(|id| {
                id.len() != 8
                    || !matches!(id.as_bytes()[0], b'0'..=b'7')
                    || !id.bytes().all(|b| b.is_ascii_hexdigit())
                    || id == "00000000"
            })
        })
    {
        return Err(Failure::refused(
            op,
            "invalidBlockId",
            "Comment paragraphs need explicit package paragraph identities.",
        ));
    }
    let mut positions = Vec::new();
    match op["anchor"]["kind"].as_str() {
        Some("point") => positions.push((
            op["anchor"]["at"].clone(),
            vec![json!({"type":"commentReference","id":id})],
        )),
        Some("range") => {
            let from = &op["anchor"]["from"];
            let to = &op["anchor"]["to"];
            if from["story"] != to["story"] {
                return Err(Failure::refused(
                    op,
                    "crossBlockRange",
                    "A comment range must stay in one story.",
                ));
            }
            let boundary = |position: &Value| {
                locate_position(document, position, op).map_err(|failure| match failure {
                    Failure::Refused { reason, .. } if reason == "blockNotFound" => {
                        Failure::refused(
                            op,
                            "blockNotFound",
                            "A comment boundary paragraph does not exist.",
                        )
                    }
                    other => other,
                })
            };
            let from_path = boundary(from)?;
            let to_path = boundary(to)?;
            let paths = main_paths(document);
            let from_index = paths
                .iter()
                .position(|path| *path == from_path)
                .expect("located paragraph");
            let to_index = paths
                .iter()
                .position(|path| *path == to_path)
                .expect("located paragraph");
            let from_gap = gap(
                content(at_path(document, &from_path).expect("located paragraph")),
                from,
                op,
            )?;
            let to_gap = gap(
                content(at_path(document, &to_path).expect("located paragraph")),
                to,
                op,
            )?;
            if from_index > to_index || from_index == to_index && from_gap > to_gap {
                return Err(Failure::refused(
                    op,
                    "invalidOffset",
                    "A comment range cannot run backwards.",
                ));
            }
            let mut end = to.clone();
            end["zeroWidthBefore"] = json!(to_gap.zero);
            let mut start = from.clone();
            start["zeroWidthBefore"] = json!(from_gap.zero);
            positions.push((
                end,
                vec![
                    json!({"type":"commentRangeEnd","id":id}),
                    json!({"type":"commentReference","id":id}),
                ],
            ));
            positions.push((start, vec![json!({"type":"commentRangeStart","id":id})]));
        }
        Some("reply") => {
            let parent = valid_id(&op["anchor"]["parentId"]);
            if !records
                .iter()
                .any(|record| valid_id(&record["id"]) == parent)
            {
                return Err(fail(op, "The parent comment does not exist."));
            }
            comment["parentId"] = op["anchor"]["parentId"].clone();
            let own: Vec<_> = anchors(document, op)?
                .into_iter()
                .filter(|anchor| anchor_id(&anchor.node) == parent)
                .collect();
            if own.is_empty() {
                return Err(fail(
                    op,
                    "The parent comment has no representable owned anchors.",
                ));
            }
            for anchor in own.into_iter().rev() {
                positions.push((json!({"story":"main","blockId":anchor.block,"offset":anchor.after_offset,"zeroWidthBefore":anchor.after_zero}),
                    vec![json!({"type":anchor.node["type"],"id":id})]));
            }
        }
        Some("revision") => return Err(Failure::unsupported(op, "revisionCommentAnchors")),
        _ => return Err(fail(op, "The comment anchor discriminator is invalid.")),
    }
    let mut after = document.clone();
    let mut updated = records;
    updated.push(comment);
    with_comments(&mut after, updated, "present", op)?;
    for (position, nodes) in positions {
        let path = locate_position(&after, &position, op)?;
        let paragraph = at_path_mut(&mut after, &path).expect("located paragraph");
        let at = gap(content(paragraph), &position, op)?;
        paragraph["content"] = json!(insert(content(paragraph), at, &nodes));
    }
    let next_rel = reconcile(&after, relations);
    finish(document, after, relations, next_rel, op, &[id])
}

fn delete(
    document: &Value,
    op: &Value,
    relations: &RelationshipState,
) -> Result<CommentApplied, Failure> {
    let records = comments(document, op)?;
    let comment = records
        .iter()
        .find(|record| record["id"] == op["id"])
        .ok_or_else(|| fail(op, "The comment does not exist."))?;
    let parent_is_comment = comment
        .get("parentId")
        .is_some_and(|id| records.iter().any(|record| record["id"] == *id));
    if !matches!(op["scope"].as_str(), Some("thread" | "reply"))
        || (op["scope"] == "reply") != parent_is_comment
    {
        return Err(fail(
            op,
            "Comment deletion scope must match root or reply ownership.",
        ));
    }
    let root_id = valid_id(&comment["id"]).expect("validated id");
    let mut ids = BTreeSet::from([root_id]);
    let mut id_order = vec![root_id];
    loop {
        let prior = ids.len();
        for record in &records {
            if record
                .get("parentId")
                .and_then(valid_id)
                .is_some_and(|id| ids.contains(&id))
            {
                let id = valid_id(&record["id"]).expect("validated id");
                if ids.insert(id) {
                    id_order.push(id);
                }
            }
        }
        if prior == ids.len() {
            break;
        }
    }
    let mut after = document.clone();
    for path in main_paths(document) {
        let paragraph = at_path_mut(&mut after, &path).expect("paragraph path");
        let seams = owned_seams(content(paragraph), &ids);
        if !seams.is_empty() {
            paragraph["content"] = json!(normalized(content(paragraph), &ids, &seams));
        }
    }
    let kept: Vec<_> = records
        .into_iter()
        .filter(|record| !ids.contains(&valid_id(&record["id"]).expect("validated id")))
        .collect();
    let presence = if kept.is_empty() { "absent" } else { "present" };
    with_comments(&mut after, kept, presence, op)?;
    let next_rel = reconcile(&after, relations);
    finish(document, after, relations, next_rel, op, &id_order)
}

fn state_array<'a>(state: &'a Value, key: &str, op: &Value) -> Result<&'a Vec<Value>, Failure> {
    state[key]
        .as_array()
        .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))
}

fn restore(
    document: &Value,
    op: &Value,
    relations: &RelationshipState,
) -> Result<CommentApplied, Failure> {
    let expected = &op["expected"];
    let state = &op["state"];
    let raw_ids = op["ids"]
        .as_array()
        .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))?;
    for state in [expected, state] {
        for key in ["relationships", "records", "anchors"] {
            state_array(state, key, op)?;
        }
        if state_array(state, "anchors", op)?
            .iter()
            .any(|entry| entry["story"] != "main")
        {
            return Err(Failure::unsupported(op, "secondaryCommentAnchors"));
        }
    }
    let ids: BTreeSet<_> = raw_ids.iter().filter_map(valid_id).collect();
    if ids.len() != raw_ids.len()
        || capture(
            document,
            relations,
            &ids,
            state_array(expected, "anchors", op)?,
        ) != *expected
    {
        return Err(Failure::refused(
            op,
            "stale",
            "The comment inverse is stale.",
        ));
    }
    let records = state_array(state, "records", op)?;
    let mut record_ids = BTreeSet::new();
    for entry in records {
        let id = valid_id(&entry["comment"]["id"]);
        if id.is_none_or(|id| !ids.contains(&id) || !record_ids.insert(id))
            || entry["index"].as_u64().is_none()
        {
            return Err(fail(
                op,
                "The restored comment records do not match the owned identities.",
            ));
        }
    }
    if !crate::identity::keys(&expected["anchors"]).is_empty()
        || !crate::identity::keys(&state["anchors"]).is_empty()
    {
        return Err(Failure::unsupported(op, "identifiedCommentScaffolding"));
    }
    if op["scaffoldIds"] != json!({"revision":[],"control":[]}) {
        return Err(fail(
            op,
            "The comment inverse scaffold identities do not match its owned cuts.",
        ));
    }
    if crate::identity::keys(&expected["anchors"]) != crate::identity::keys(&state["anchors"]) {
        return Err(fail(
            op,
            "The comment inverse changes an unowned identified record.",
        ));
    }
    let next_anchors = state_array(state, "anchors", op)?;
    let prior_anchors = state_array(expected, "anchors", op)?;
    if next_anchors.len() != prior_anchors.len() {
        return Err(fail(
            op,
            "A comment inverse cannot replace text or unowned anchors.",
        ));
    }
    let mut after = document.clone();
    for (entry, prior) in next_anchors.iter().zip(prior_anchors) {
        let nodes = entry["content"]
            .as_array()
            .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))?;
        let old_nodes = prior["content"]
            .as_array()
            .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))?;
        validate_content(nodes, op)?;
        validate_content(old_nodes, op)?;
        let seams = owned_seams(nodes, &ids)
            .union(&owned_seams(old_nodes, &ids))
            .copied()
            .collect();
        let unowned = |nodes: &[Value]| {
            nodes
                .iter()
                .filter(|node| anchor_id(node).is_some_and(|id| !ids.contains(&id)))
                .cloned()
                .collect::<Vec<_>>()
        };
        if entry["story"] != prior["story"]
            || !entry["blockId"]
                .as_str()
                .zip(prior["blockId"].as_str())
                .is_some_and(|(a, b)| a.eq_ignore_ascii_case(b))
            || normalized(nodes, &ids, &seams) != normalized(old_nodes, &ids, &seams)
            || unowned(nodes) != unowned(old_nodes)
        {
            return Err(fail(
                op,
                "A comment inverse cannot replace text or unowned anchors.",
            ));
        }
        let path = locate_position(&after, entry, op).map_err(|failure| match failure {
            Failure::Unsupported { .. } => failure,
            _ => Failure::refused(op, "stale", "An owned comment paragraph is missing."),
        })?;
        at_path_mut(&mut after, &path).expect("located paragraph")["content"] =
            entry["content"].clone();
    }
    let mut remaining: Vec<_> = comments(document, op)?
        .into_iter()
        .filter(|record| !valid_id(&record["id"]).is_some_and(|id| ids.contains(&id)))
        .collect();
    for entry in records {
        let index = entry["index"]
            .as_u64()
            .and_then(|n| usize::try_from(n).ok())
            .filter(|n| *n <= remaining.len())
            .ok_or_else(|| {
                Failure::refused(
                    op,
                    "stale",
                    "A restored comment declaration position is stale.",
                )
            })?;
        remaining.insert(index, entry["comment"].clone());
    }
    let presence = state["listPresence"]
        .as_str()
        .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))?;
    if presence != "present" && !remaining.is_empty() {
        return Err(fail(
            op,
            "A missing comment list cannot discard unowned definitions.",
        ));
    }
    with_comments(&mut after, remaining, presence, op)?;
    let mut next_rel: Vec<_> = relations
        .entries()
        .iter()
        .filter(|(_, relation)| !is_relation(relation))
        .cloned()
        .collect();
    for entry in state_array(state, "relationships", op)? {
        let key = entry["key"].as_str().ok_or_else(|| {
            fail(
                op,
                "The comment inverse cannot replace an unowned relationship.",
            )
        })?;
        let relation = &entry["relationship"];
        if state_array(expected, "relationships", op)?
            .iter()
            .find(|old| old["key"] == key)
            .is_some_and(|old| old["relationship"] != *relation)
        {
            return Err(fail(
                op,
                "A comment inverse cannot alter an existing relationship payload.",
            ));
        }
        let index = entry["index"]
            .as_u64()
            .and_then(|n| usize::try_from(n).ok())
            .filter(|n| *n <= next_rel.len());
        if !is_relation(relation)
            || relation["id"] != key
            || index.is_none()
            || next_rel.iter().any(|(old, _)| old == key)
        {
            return Err(fail(
                op,
                "The comment inverse cannot replace an unowned relationship.",
            ));
        }
        next_rel.insert(
            index.expect("validated index"),
            (key.into(), relation.clone()),
        );
    }
    let rel_presence = state["relationshipPresence"]
        .as_str()
        .ok_or_else(|| fail(op, "The comment inverse payload is invalid."))?;
    if rel_presence != "present" && !next_rel.is_empty() {
        return Err(fail(
            op,
            "A missing relationship map cannot discard unowned relationships.",
        ));
    }
    let restored = match rel_presence {
        "present" => RelationshipState::Present(next_rel),
        "absent" => RelationshipState::Absent,
        "undefined" => RelationshipState::Undefined,
        _ => return Err(fail(op, "The comment inverse payload is invalid.")),
    };
    let id_order: Vec<_> = raw_ids.iter().filter_map(valid_id).collect();
    finish(document, after, relations, restored, op, &id_order)
}

/// Semantic adapter: the document must omit relationships, supplied losslessly here.
pub fn edit_with_relationships(
    document: &Value,
    op: &Value,
    relationships: RelationshipState,
) -> Result<CommentApplied, Failure> {
    if document.pointer("/package/relationships").is_some() {
        return Err(Failure::unsupported(op, "documentMapTransport"));
    }
    if op.get("undefinedFields").is_some() {
        return Err(Failure::unsupported(op, "commentOwnUndefined"));
    }
    if op["type"] == "restoreCommentState" {
        return restore(document, op, &relationships);
    }
    issue(document, op)?;
    match op["type"].as_str() {
        Some("createComment") => create(document, op, &relationships),
        Some("deleteComment") => delete(document, op, &relationships),
        _ => Err(Failure::unsupported(op, "commentOperationKind")),
    }
}

/// JSON alone cannot carry the JS Map created by relationship reconciliation.
pub fn edit(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let result = edit_with_relationships(document, op, RelationshipState::Absent)?;
    if result.relationships != RelationshipState::Absent {
        return Err(Failure::unsupported(op, "documentMapTransport"));
    }
    Ok(result.applied)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn document() -> Value {
        json!({"package":{"document":{"type":"document","content":[
            {"type":"paragraph","paraId":"00000001","metadata":{"keep":true},"content":[
                {"type":"run","formatting":{"bold":true},"authored":"run","content":[{"type":"text","text":"abcédef","authored":"text"}]}]},
            {"type":"paragraph","paraId":"00000002","content":[]}],"unknownBody":17},"unknownPackage":{"keep":[1,2,3]}},"unknownDocument":"keep"})
    }
    fn create_op(from: usize, to: usize) -> Value {
        json!({"type":"createComment","comment":{"id":3,"author":"author","unknownComment":true,"content":[
            {"type":"paragraph","paraId":"00000003","content":[{"type":"run","content":[{"type":"text","text":"comment"}]}]}]},
            "anchor":{"kind":"range","from":{"story":"main","blockId":"00000001","offset":from},"to":{"story":"main","blockId":"00000001","offset":to}}})
    }
    #[test]
    fn every_scalar_range_restores_exact_metadata_and_relationship_presence() {
        for from in 0..=7 {
            for to in from..=7 {
                let before = document();
                for relationships in [
                    RelationshipState::Absent,
                    RelationshipState::Undefined,
                    RelationshipState::Present(vec![(
                        "rId8".into(),
                        json!({"id":"rId8","type":"unknown","target":"keep","metadata":true}),
                    )]),
                ] {
                    let changed = edit_with_relationships(
                        &before,
                        &create_op(from, to),
                        relationships.clone(),
                    )
                    .unwrap();
                    assert_eq!(changed.applied.touched.modified, vec!["00000001"]);
                    assert_eq!(
                        changed.applied.document["package"]["unknownPackage"],
                        before["package"]["unknownPackage"]
                    );
                    assert_eq!(
                        changed.applied.inverse[0]["scaffoldIds"],
                        json!({"revision":[],"control":[]})
                    );
                    let restored = edit_with_relationships(
                        &changed.applied.document,
                        &changed.applied.inverse[0],
                        changed.relationships,
                    )
                    .unwrap();
                    assert_eq!(restored.applied.document, before);
                    assert_eq!(restored.relationships, relationships);
                }
            }
        }
    }
    #[test]
    fn reply_thread_deletion_and_redo_preserve_owned_anchor_closure() {
        let before = document();
        let mut root_op = create_op(1, 5);
        root_op["comment"]["id"] = json!(9);
        let root = edit_with_relationships(&before, &root_op, RelationshipState::Absent).unwrap();
        let mut reply = create_op(0, 0);
        reply["comment"]["content"][0]["paraId"] = json!("00000004");
        reply["anchor"] = json!({"kind":"reply","parentId":9});
        let replied =
            edit_with_relationships(&root.applied.document, &reply, root.relationships).unwrap();
        let deleted = edit_with_relationships(
            &replied.applied.document,
            &json!({"type":"deleteComment","id":9,"scope":"thread"}),
            replied.relationships.clone(),
        )
        .unwrap();
        assert!(
            deleted
                .applied
                .document
                .pointer("/package/document/comments")
                .is_none()
        );
        assert_eq!(deleted.applied.inverse[0]["ids"], json!([9, 3]));
        let restored = edit_with_relationships(
            &deleted.applied.document,
            &deleted.applied.inverse[0],
            deleted.relationships,
        )
        .unwrap();
        assert_eq!(restored.applied.document, replied.applied.document);
        assert_eq!(restored.relationships, replied.relationships);
        let redone = edit_with_relationships(
            &restored.applied.document,
            &restored.applied.inverse[0],
            restored.relationships,
        )
        .unwrap();
        assert_eq!(redone.applied.document, deleted.applied.document);
    }
    #[test]
    fn inverse_cannot_change_unowned_text_and_json_reports_map_transport() {
        let before = document();
        assert!(
            matches!(edit(&before,&create_op(1,4)),Err(Failure::Unsupported{dimension,..}) if dimension=="documentMapTransport")
        );
        let changed =
            edit_with_relationships(&before, &create_op(1, 4), RelationshipState::Absent).unwrap();
        let mut forged = changed.applied.inverse[0].clone();
        forged["state"]["anchors"][0]["content"][0]["content"][0]["text"] = json!("forged");
        assert!(
            matches!(edit_with_relationships(&changed.applied.document,&forged,changed.relationships),Err(Failure::Refused{reason,..}) if reason=="structureMismatch")
        );
    }
    #[test]
    fn stale_scope_collisions_and_absent_map_delete_closure() {
        let before = document();
        let changed =
            edit_with_relationships(&before, &create_op(1, 4), RelationshipState::Absent).unwrap();
        assert!(
            matches!(edit_with_relationships(&changed.applied.document,&create_op(1,4),changed.relationships.clone()),Err(Failure::Refused{reason,..}) if reason=="idCollision")
        );
        assert!(
            matches!(edit_with_relationships(&changed.applied.document,&json!({"type":"deleteComment","id":3,"scope":"reply"}),changed.relationships.clone()),Err(Failure::Refused{reason,..}) if reason=="structureMismatch")
        );
        let mut stale = changed.applied.inverse[0].clone();
        stale["expected"]["records"][0]["comment"]["author"] = json!("stale");
        assert!(
            matches!(edit_with_relationships(&changed.applied.document,&stale,changed.relationships),Err(Failure::Refused{reason,..}) if reason=="stale")
        );
        // Loaded documents may have no relationship property. JSON can represent
        // this deletion and its exact restore without pretending to transport Map.
        let deleted = edit(
            &changed.applied.document,
            &json!({"type":"deleteComment","id":3,"scope":"thread"}),
        )
        .unwrap();
        let restored = edit(&deleted.document, &deleted.inverse[0]).unwrap();
        assert_eq!(restored.document, changed.applied.document);
    }
    #[test]
    fn span_across_blocks_and_point_in_empty_paragraph_restore() {
        let before = document();
        let mut op = create_op(2, 0);
        op["anchor"]["to"]["blockId"] = json!("00000002");
        let changed = edit_with_relationships(&before, &op, RelationshipState::Absent).unwrap();
        assert_eq!(
            changed.applied.touched.modified,
            vec!["00000001", "00000002"]
        );
        let restored = edit_with_relationships(
            &changed.applied.document,
            &changed.applied.inverse[0],
            changed.relationships,
        )
        .unwrap();
        assert_eq!(restored.applied.document, before);
        op["anchor"] =
            json!({"kind":"point","at":{"story":"main","blockId":"00000002","offset":0}});
        let point = edit_with_relationships(&before, &op, RelationshipState::Absent).unwrap();
        assert_eq!(
            point.applied.document["package"]["document"]["content"][1]["content"],
            json!([{"type":"commentReference","id":3}])
        );
    }
}
