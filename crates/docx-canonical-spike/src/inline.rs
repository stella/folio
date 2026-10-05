//! Plain-run subset of TS `inline.ts` and `leaves.ts`.
//!
//! Partitioning clones each authored record only along the cut. Open slice
//! depths describe those actual run/text boundaries; deletion never flattens
//! adjacent authored runs. Unsupported record kinds are reported explicitly.

use crate::refusal::Failure;
use serde_json::{Map, Value, json};

fn unsupported(op: &Value, dimension: &str) -> Failure {
    Failure::unsupported(op, dimension)
}

fn refused(op: &Value, reason: &str, message: &str) -> Failure {
    Failure::refused(op, reason, message)
}

fn content(value: &Value) -> &[Value] {
    value["content"]
        .as_array()
        .expect("validated content array")
}

fn text(value: &Value) -> &str {
    value["text"].as_str().expect("validated text leaf")
}

fn with_content(value: &Value, items: Vec<Value>) -> Value {
    let mut result = value.clone();
    result["content"] = Value::Array(items);
    result
}

fn with_text(value: &Value, replacement: String) -> Value {
    let mut result = value.clone();
    result["text"] = Value::String(replacement);
    result
}

fn width(value: &Value) -> usize {
    if value["type"] == "text" {
        text(value).encode_utf16().count()
    } else {
        content(value).iter().map(width).sum()
    }
}

pub(crate) fn list_width(items: &[Value]) -> usize {
    items.iter().map(width).sum()
}

pub(crate) fn validate_plain(items: &[Value], op: &Value, incoming: bool) -> Result<(), Failure> {
    for run in items {
        if run.get("type").and_then(Value::as_str) != Some("run") {
            return Err(unsupported(op, "inlineContainersAndSpecialLeaves"));
        }
        if run.get("propertyChanges").is_some() {
            return Err(unsupported(op, "runPropertyRevisions"));
        }
        let children = run
            .get("content")
            .and_then(Value::as_array)
            .ok_or_else(|| unsupported(op, "runContentShape"))?;
        if children.is_empty() {
            return Err(if incoming {
                refused(op, "emptyContent", "The slice holds an empty run.")
            } else {
                unsupported(op, "emptyAuthoredRecords")
            });
        }
        for child in children {
            if child.get("type").and_then(Value::as_str) != Some("text") {
                return Err(unsupported(op, "inlineContainersAndSpecialLeaves"));
            }
            let value = child
                .get("text")
                .and_then(Value::as_str)
                .ok_or_else(|| unsupported(op, "textLeafShape"))?;
            if value.is_empty() {
                return Err(if incoming {
                    refused(op, "emptyContent", "The slice holds an empty text node.")
                } else {
                    unsupported(op, "emptyAuthoredRecords")
                });
            }
            if incoming && illegal_xml(value) {
                return Err(refused(
                    op,
                    "invalidText",
                    "The slice holds text that cannot be written.",
                ));
            }
        }
    }
    Ok(())
}

pub(crate) fn illegal_xml(value: &str) -> bool {
    value.chars().any(|character| {
        !matches!(character as u32,
        0x9 | 0xa | 0xd | 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
    })
}

/// The byte boundary for a UTF-16 offset; a surrogate-pair interior has none.
fn byte_at(value: &str, offset: usize) -> Option<usize> {
    let mut units = 0;
    for (byte, character) in value.char_indices() {
        if units == offset {
            return Some(byte);
        }
        units += character.len_utf16();
        if units > offset {
            return None;
        }
    }
    (units == offset).then_some(value.len())
}

pub(crate) fn coordinate(items: &[Value], position: &Value, op: &Value) -> Result<usize, Failure> {
    let offset = position
        .get("offset")
        .and_then(Value::as_u64)
        .and_then(|number| usize::try_from(number).ok())
        .filter(|offset| *offset <= list_width(items))
        .ok_or_else(|| {
            refused(
                op,
                "invalidOffset",
                "The offset is not a position in the paragraph.",
            )
        })?;
    if position
        .get("zeroWidthBefore")
        .is_some_and(|value| value.as_u64() != Some(0))
    {
        return Err(refused(
            op,
            "invalidOffset",
            "The zero-width coordinate is unavailable.",
        ));
    }
    let mut start = 0;
    for run in items {
        for child in content(run) {
            let end = start + width(child);
            if start <= offset && offset <= end && byte_at(text(child), offset - start).is_none() {
                return Err(refused(
                    op,
                    "splitsSurrogatePair",
                    "The position splits a surrogate pair.",
                ));
            }
            start = end;
        }
    }
    Ok(offset)
}

fn count(op: &Value, value: Option<&Value>, message: &str) -> Result<usize, Failure> {
    value
        .and_then(Value::as_u64)
        .and_then(|number| usize::try_from(number).ok())
        .ok_or_else(|| refused(op, "structureMismatch", message))
}

pub(crate) fn seam_depth(op: &Value) -> Result<usize, Failure> {
    let action = if op["type"] == "splitInline" {
        "split"
    } else {
        "join"
    };
    let message = format!(
        "A {action} {} one level or more.",
        if action == "split" { "cuts" } else { "merges" }
    );
    let depth = count(op, op.get("depth"), &message)?;
    if depth == 0 {
        return Err(refused(op, "structureMismatch", &message));
    }
    Ok(depth)
}

fn position(source: &Value, paragraph: &Value, offset: usize) -> Value {
    json!({"story":source["story"],"blockId":paragraph["paraId"],"offset":offset,"zeroWidthBefore":0})
}

struct Pieces {
    pieces: Vec<(usize, Value)>,
    min: usize,
    max: usize,
}

fn region(gaps: &[usize], unit: usize) -> usize {
    gaps.iter().filter(|gap| **gap <= unit).count()
}

/// Port of partitionNode for the two supported record kinds.
fn partition_node(node: &Value, gaps: &[usize], cursor: &mut usize) -> Pieces {
    if node["type"] == "text" {
        let start = *cursor;
        let value = text(node);
        let length = width(node);
        let mut cuts = vec![0];
        cuts.extend(
            gaps.iter()
                .copied()
                .filter(|gap| start < *gap && *gap < start + length)
                .map(|gap| gap - start),
        );
        cuts.push(length);
        cuts.dedup();
        let pieces = cuts
            .windows(2)
            .map(|range| {
                let from = byte_at(value, range[0]).expect("validated scalar boundary");
                let to = byte_at(value, range[1]).expect("validated scalar boundary");
                let piece = if cuts.len() == 2 {
                    node.clone()
                } else {
                    with_text(node, value[from..to].to_owned())
                };
                (region(gaps, start + range[0]), piece)
            })
            .collect::<Vec<_>>();
        *cursor += length;
        return Pieces {
            min: pieces[0].0,
            max: pieces[pieces.len() - 1].0,
            pieces,
        };
    }
    let mut regions = vec![Vec::new(); gaps.len() + 1];
    let mut min = usize::MAX;
    let mut max = 0;
    for child in content(node) {
        let inner = partition_node(child, gaps, cursor);
        min = min.min(inner.min);
        max = max.max(inner.max);
        for (region, piece) in inner.pieces {
            regions[region].push(piece);
        }
    }
    if min == max {
        return Pieces {
            pieces: vec![(min, node.clone())],
            min,
            max,
        };
    }
    let pieces = regions
        .into_iter()
        .enumerate()
        .filter(|(_, items)| !items.is_empty())
        .map(|(region, items)| (region, with_content(node, items)))
        .collect();
    Pieces { pieces, min, max }
}

pub(crate) fn partition(items: &[Value], gaps: &[usize]) -> Vec<Vec<Value>> {
    let mut regions = vec![Vec::new(); gaps.len() + 1];
    let mut cursor = 0;
    for item in items {
        for (region, piece) in partition_node(item, gaps, &mut cursor).pieces {
            regions[region].push(piece);
        }
    }
    regions
}

/// Port of spanningRecords: the run and text node whose leaves cross the cut.
pub(crate) fn spanning(items: &[Value], gaps: &[usize], first: usize, last: usize) -> usize {
    let mut cursor = 0;
    for item in items {
        let start = cursor;
        let pieces = partition_node(item, gaps, &mut cursor);
        if pieces.min > first || pieces.max < last {
            continue;
        }
        let mut child_cursor = start;
        for child in content(item) {
            let inner = partition_node(child, gaps, &mut child_cursor);
            if inner.min <= first && inner.max >= last {
                return 2;
            }
        }
        return 1;
    }
    0
}

fn same_run_fields(left: &Value, right: &Value) -> bool {
    let fields = |value: &Value| {
        value
            .as_object()
            .expect("validated run")
            .iter()
            .filter(|(key, _)| key.as_str() != "content")
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect::<Map<_, _>>()
    };
    fields(left) == fields(right)
}

/// Port of exact mergeLists. A text merge spreads the left text record, even
/// when the enclosing open run continues the right source record.
pub(crate) fn merge(
    left: &[Value],
    right: &[Value],
    depth: usize,
    source_second: bool,
) -> Option<Vec<Value>> {
    if depth == 0 {
        return Some(left.iter().chain(right).cloned().collect());
    }
    let last = left.last()?;
    let first = right.first()?;
    let merged = if last["type"] == "text" && first["type"] == "text" && depth == 1 {
        with_text(last, format!("{}{}", text(last), text(first)))
    } else if last["type"] == "run" && first["type"] == "run" && same_run_fields(last, first) {
        let children = merge(content(last), content(first), depth - 1, source_second)?;
        with_content(if source_second { first } else { last }, children)
    } else {
        return None;
    };
    let mut result = left[..left.len() - 1].to_vec();
    result.push(merged);
    result.extend_from_slice(&right[1..]);
    Some(result)
}

pub(crate) fn delete_between(items: &[Value], from: usize, to: usize) -> (Vec<Value>, Value) {
    let regions = partition(items, &[from, to]);
    let across = spanning(items, &[from, to], 0, 2);
    let remaining = merge(&regions[0], &regions[2], across, false)
        .expect("two pieces of the same cut record merge");
    let removed = json!({"content":regions[1],
        "openStart":spanning(items, &[from,to],0,1),
        "openEnd":spanning(items, &[from,to],1,2)});
    (remaining, removed)
}

fn insertion_inverse(
    before: &[Value],
    after: &[Value],
    start: usize,
    end: usize,
) -> Option<(Value, usize)> {
    let (restored, removed) = delete_between(after, start, end);
    let cut = spanning(before, &[start], 0, 1);
    let across = spanning(&restored, &[start], 0, 1);
    (across <= cut).then(|| (removed, cut - across))
}

fn new_text_run(value: &str, formatting: Option<&Value>) -> Value {
    let mut run = json!({"type":"run","content":[{"type":"text","text":value}]});
    if let Some(formatting) =
        formatting.filter(|value| value.as_object().is_some_and(|fields| !fields.is_empty()))
    {
        run["formatting"] = formatting.clone();
    }
    run
}

fn run_with_text(run: &Value, offset: usize, inserted: &str) -> Value {
    let mut cursor = 0;
    let mut children = content(run).to_vec();
    for (index, child) in content(run).iter().enumerate() {
        let end = cursor + width(child);
        if cursor <= offset && offset <= end {
            let value = text(child);
            let byte = byte_at(value, offset - cursor).expect("validated scalar boundary");
            children[index] = with_text(
                child,
                format!("{}{}{}", &value[..byte], inserted, &value[byte..]),
            );
            return with_content(run, children);
        }
        cursor = end;
    }
    unreachable!("nonempty insertion host contains the position")
}

fn insert_text(items: &[Value], offset: usize, inserted: &str, run_props: &Value) -> Vec<Value> {
    let formatting = (run_props != "inherit").then_some(run_props);
    if items.is_empty() {
        return vec![new_text_run(inserted, formatting)];
    }
    let host_unit = if offset > 0 { offset - 1 } else { 0 };
    let mut cursor = 0;
    for (index, run) in items.iter().enumerate() {
        let end = cursor + width(run);
        if cursor <= host_unit && host_unit < end {
            let local = offset - cursor;
            let empty = json!({});
            let same_formatting = formatting
                .is_none_or(|incoming| incoming == run.get("formatting").unwrap_or(&empty));
            let mut result = items[..index].to_vec();
            if same_formatting {
                result.push(run_with_text(run, local, inserted));
            } else {
                let halves = partition(content(run), &[local]);
                if !halves[0].is_empty() {
                    result.push(with_content(run, halves[0].clone()));
                }
                result.push(new_text_run(inserted, formatting));
                if !halves[1].is_empty() {
                    result.push(with_content(run, halves[1].clone()));
                }
            }
            result.extend_from_slice(&items[index + 1..]);
            return result;
        }
        cursor = end;
    }
    unreachable!("validated offset has an insertion host")
}

fn inserted_result(
    paragraph: &Value,
    op: &Value,
    before: &[Value],
    after: Vec<Value>,
    start: usize,
    end: usize,
) -> Result<(Value, Vec<Value>), Failure> {
    let (removed, join) = insertion_inverse(before, &after, start, end).ok_or_else(|| {
        refused(
            op,
            "structureMismatch",
            "The insertion would merge records that were separate.",
        )
    })?;
    let mut inverse = json!({"type":"deleteRange","from":position(&op["at"],paragraph,start),
        "to":position(&op["at"],paragraph,end),"expected":removed});
    if join > 0 {
        inverse["join"] = json!(join);
    }
    Ok((with_content(paragraph, after), vec![inverse]))
}

pub fn edit(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    if op.get("revision").is_some() {
        return Err(unsupported(op, "trackedInlineEdits"));
    }
    if op
        .get("undefinedFields")
        .and_then(Value::as_array)
        .is_some_and(|paths| !paths.is_empty())
    {
        return Err(unsupported(op, "inlineOwnUndefinedFields"));
    }
    let items = paragraph
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| unsupported(op, "paragraphContentShape"))?;
    validate_plain(items, op, false)?;
    match op.get("type").and_then(Value::as_str) {
        Some("splitInline" | "joinInline") => {
            let depth = seam_depth(op)?;
            let offset = coordinate(items, &op["at"], op).map_err(|failure| match failure {
                Failure::Refused { reason, .. } => refused(
                    op,
                    &reason,
                    &format!(
                        "{} is not a position in {}.",
                        op["at"]["offset"],
                        op["at"]["blockId"].as_str().unwrap_or_default()
                    ),
                ),
                failure => failure,
            })?;
            let across = spanning(items, &[offset], 0, 1);
            let halves = partition(items, &[offset]);
            let split = op["type"] == "splitInline";
            let after = if split {
                if depth > across {
                    return Err(refused(
                        op,
                        "structureMismatch",
                        &format!(
                            "Fewer than {depth} records run across {offset} in {}.",
                            op["at"]["blockId"].as_str().unwrap_or_default()
                        ),
                    ));
                }
                merge(&halves[0], &halves[1], across - depth, false)
                    .expect("two pieces of the same cut record merge")
            } else {
                across
                    .checked_add(depth)
                    .and_then(|levels| merge(&halves[0], &halves[1], levels, false))
                    .ok_or_else(|| {
                        refused(
                            op,
                            "structureMismatch",
                            &format!(
                                "The records meeting at {offset} in {} cannot be merged.",
                                op["at"]["blockId"].as_str().unwrap_or_default()
                            ),
                        )
                    })?
            };
            let inverse = json!({"type":if split {"joinInline"} else {"splitInline"},
                "at":position(&op["at"], paragraph, offset),"depth":depth});
            Ok((with_content(paragraph, after), vec![inverse]))
        }
        Some("insertText") => {
            let inserted = op.get("text").and_then(Value::as_str).ok_or_else(|| {
                refused(op, "invalidOperation", "An insertion must name its text.")
            })?;
            if inserted.is_empty() || inserted.contains(['\t', '\n', '\r']) || illegal_xml(inserted)
            {
                return Err(refused(op, "invalidText", "The text cannot be inserted."));
            }
            let offset = coordinate(items, &op["at"], op)?;
            let props = op
                .get("runProps")
                .filter(|props| **props == "inherit" || props.is_object())
                .ok_or_else(|| {
                    refused(
                        op,
                        "invalidOperation",
                        "An insertion must name its run properties.",
                    )
                })?;
            let after = insert_text(items, offset, inserted, props);
            inserted_result(
                paragraph,
                op,
                items,
                after,
                offset,
                offset + inserted.encode_utf16().count(),
            )
        }
        Some("insertContent") => {
            let slice = op.get("slice").and_then(Value::as_object).ok_or_else(|| {
                refused(op, "invalidOperation", "An insertion must name its slice.")
            })?;
            let incoming = slice
                .get("content")
                .and_then(Value::as_array)
                .ok_or_else(|| {
                    refused(
                        op,
                        "invalidOperation",
                        "A slice must contain a content array.",
                    )
                })?;
            if incoming.is_empty() {
                return Err(refused(op, "emptyContent", "The slice holds nothing."));
            }
            validate_plain(incoming, op, true)?;
            let open_start = count(op, slice.get("openStart"), "Open depths are counts.")?;
            let open_end = count(op, slice.get("openEnd"), "Open depths are counts.")?;
            let offset = coordinate(items, &op["at"], op)?;
            let halves = partition(items, &[offset]);
            let start = merge(&halves[0], incoming, open_start, false);
            let after = start
                .and_then(|start| merge(&start, &halves[1], open_end, true))
                .ok_or_else(|| {
                    refused(
                        op,
                        "structureMismatch",
                        "The slice's open ends do not fit the position.",
                    )
                })?;
            inserted_result(
                paragraph,
                op,
                items,
                after,
                offset,
                offset + list_width(incoming),
            )
        }
        Some("deleteRange") => {
            let join = op
                .get("join")
                .map(|value| count(op, Some(value), "A join depth is a count."))
                .transpose()?
                .unwrap_or(0);
            let from = coordinate(items, &op["from"], op)?;
            let to = coordinate(items, &op["to"], op)?;
            if from > to {
                return Err(refused(op, "invalidOffset", "The range runs backwards."));
            }
            if from == to {
                if op.get("expected").is_none_or(|expected| {
                    expected
                        .get("content")
                        .and_then(Value::as_array)
                        .is_some_and(Vec::is_empty)
                }) {
                    return Ok((paragraph.clone(), vec![]));
                }
                return Err(refused(op, "stale", "The range to delete is empty."));
            }
            let (mut after, removed) = delete_between(items, from, to);
            if op
                .get("expected")
                .is_some_and(|expected| *expected != removed)
            {
                return Err(refused(
                    op,
                    "stale",
                    "The range holds other content than expected.",
                ));
            }
            if join > 0 {
                let depth = spanning(&after, &[from], 0, 1) + join;
                let halves = partition(&after, &[from]);
                after = merge(&halves[0], &halves[1], depth, false).ok_or_else(|| {
                    refused(
                        op,
                        "structureMismatch",
                        "The records meeting at the gap cannot be merged.",
                    )
                })?;
            }
            let inverse = json!({"type":"insertContent","at":position(&op["from"],paragraph,from),"slice":removed});
            Ok((with_content(paragraph, after), vec![inverse]))
        }
        _ => Err(unsupported(op, "inlineOperationKind")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph() -> Value {
        json!({"type":"paragraph","paraId":"00000001","content":[
            {"type":"run","preservedAttributes":[{"name":"rsidR","value":"1234"}],"content":[
                {"type":"text","text":"a😀b","extra":"left"}, {"type":"text","text":"cd","extra":"right"}]},
            {"type":"run","formatting":{"bold":true},"content":[{"type":"text","text":"EF"}]}]})
    }

    fn at(offset: usize) -> Value {
        json!({"story":"main","blockId":"00000001","offset":offset})
    }

    #[test]
    fn every_scalar_range_roundtrips_authored_boundaries_and_metadata() {
        let original = paragraph();
        let boundaries = [0, 1, 3, 4, 5, 6, 7, 8];
        for from in boundaries {
            for to in boundaries.into_iter().filter(|to| *to > from) {
                let op = json!({"type":"deleteRange","from":at(from),"to":at(to)});
                let (changed, inverses) = edit(&original, &op).unwrap();
                let (restored, _) = edit(&changed, &inverses[0]).unwrap();
                assert_eq!(restored, original, "range {from}..{to}");
            }
        }
    }

    #[test]
    fn every_scalar_insertion_roundtrips_inherited_and_explicit_formatting() {
        let original = paragraph();
        for offset in [0, 1, 3, 4, 5, 6, 7, 8] {
            for props in [json!("inherit"), json!({}), json!({"italic":true})] {
                let op = json!({"type":"insertText","at":at(offset),"text":"ž𐐀","runProps":props});
                let (changed, inverses) = edit(&original, &op).unwrap();
                let (restored, _) = edit(&changed, &inverses[0]).unwrap();
                assert_eq!(restored, original, "offset {offset}, props {props}");
            }
        }
    }

    #[test]
    fn every_available_scalar_split_restores_authored_records() {
        let original = paragraph();
        for offset in [0, 1, 3, 4, 5, 6, 7, 8] {
            let across = spanning(content(&original), &[offset], 0, 1);
            for depth in 1..=across {
                let op = json!({"type":"splitInline","at":at(offset),"depth":depth});
                let (changed, inverse) = edit(&original, &op).unwrap();
                assert_eq!(
                    inverse,
                    vec![json!({"type":"joinInline","at":{
                    "story":"main","blockId":"00000001","offset":offset,"zeroWidthBefore":0
                },"depth":depth})]
                );
                let (restored, redo) = edit(&changed, &inverse[0]).unwrap();
                assert_eq!(restored, original, "offset {offset}, depth {depth}");
                assert_eq!(edit(&restored, &redo[0]).unwrap().0, changed);
            }
        }
    }

    #[test]
    fn rejects_surrogate_pair_interior_and_refuses_stale_slice() {
        let original = paragraph();
        let op = json!({"type":"insertText","at":at(2),"text":"x","runProps":"inherit"});
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "splitsSurrogatePair")
        );
        let op = json!({"type":"deleteRange","from":at(0),"to":at(1),"expected":{"content":[],"openStart":0,"openEnd":0}});
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "stale")
        );
    }
}
