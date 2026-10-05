//! Tracked text edits over plain runs, plus removal of whole insertion wrappers.
//! Existing wrappers, atoms, and property reviews are refused as unsupported
//! dimensions rather than approximated by flattening their authored structure.

use crate::{inline, refusal::Failure};
use serde_json::{Map, Value, json};

fn items<'a>(paragraph: &'a Value, op: &Value) -> Result<&'a [Value], Failure> {
    paragraph
        .get("content")
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .ok_or_else(|| Failure::unsupported(op, "paragraphContentShape"))
}

fn with_content(paragraph: &Value, content: Vec<Value>) -> Value {
    let mut result = paragraph.clone();
    result["content"] = Value::Array(content);
    result
}

fn position(source: &Value, paragraph: &Value, offset: usize) -> Value {
    json!({"story":source["story"],"blockId":paragraph["paraId"],
        "offset":offset,"zeroWidthBefore":0})
}

fn restoring(paragraph: &Value, changed: &Value, source: &Value) -> Value {
    json!({"type":"replaceInline","story":source["story"],
        "blockId":paragraph["paraId"],"expected":changed["content"],
        "content":paragraph["content"]})
}

fn stamp_info(op: &Value) -> Result<Value, Failure> {
    let stamp = op
        .get("revision")
        .and_then(Value::as_object)
        .ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOperation",
                "A tracked edit must name its revision.",
            )
        })?;
    let mut info = Map::new();
    for key in ["id", "author", "date"] {
        let value = stamp.get(key).ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOperation",
                "A revision has incomplete metadata.",
            )
        })?;
        info.insert(key.to_owned(), value.clone());
    }
    if let Some(initials) = stamp.get("initials") {
        info.insert("initials".to_owned(), initials.clone());
    }
    Ok(Value::Object(info))
}

/// `wrapList`/`groupCovered` for plain run/text source records. Plain runs
/// have no identity slots, so the deletion's retainedAfter list is absent.
fn wrap_plain(
    content: &[Value],
    from: usize,
    to: usize,
    kind: &str,
    info: Value,
    resolution_join: usize,
) -> Vec<Value> {
    let regions = inline::partition(content, &[from, to]);
    let wrapper = json!({"type":kind,"info":info,"content":regions[1],
        "resolutionJoins":{
            "before":inline::spanning(content,&[from],0,1),
            "after":inline::spanning(content,&[to],0,1),
            "remove":resolution_join}});
    let mut result = regions[0].clone();
    result.push(wrapper);
    result.extend_from_slice(&regions[2]);
    result
}

/// Validate scalar offsets through whole insertion wrappers, retaining those
/// wrappers in the actual edit. The flattened list is only a coordinate view.
fn coordinate_view(content: &[Value], op: &Value, incoming: bool) -> Result<Vec<Value>, Failure> {
    let mut view = Vec::new();
    for node in content {
        match node.get("type").and_then(Value::as_str) {
            Some("run") => {
                inline::validate_plain(std::slice::from_ref(node), op, incoming)?;
                view.push(node.clone());
            }
            Some("insertion") => {
                let children = node
                    .get("content")
                    .and_then(Value::as_array)
                    .ok_or_else(|| Failure::unsupported(op, "trackedWrapperContentShape"))?;
                if children.is_empty() {
                    return Err(if incoming {
                        Failure::refused(
                            op,
                            "emptyContent",
                            "The slice holds an empty revision wrapper.",
                        )
                    } else {
                        Failure::unsupported(op, "emptyAuthoredRecords")
                    });
                }
                inline::validate_plain(children, op, incoming)?;
                view.extend_from_slice(children);
            }
            _ => {
                return Err(Failure::unsupported(
                    op,
                    "trackedInlineContainersAndSpecialLeaves",
                ));
            }
        }
    }
    Ok(view)
}

/// Close the inverse of a tracked insertion without supporting arbitrary
/// wrapper repartitioning. No wrapper cut means no source-slot provenance
/// rewrite or revision-id allocation is needed.
fn delete_whole_insertions(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    let content = items(paragraph, op)?;
    let view = coordinate_view(content, op, false)?;
    let join = op
        .get("join")
        .map(|value| {
            value
                .as_u64()
                .and_then(|number| usize::try_from(number).ok())
                .ok_or_else(|| {
                    Failure::refused(op, "structureMismatch", "A join depth is a count.")
                })
        })
        .transpose()?
        .unwrap_or(0);
    let from = inline::coordinate(&view, &op["from"], op)?;
    let to = inline::coordinate(&view, &op["to"], op)?;
    if from > to {
        return Err(Failure::refused(
            op,
            "invalidOffset",
            "The range runs backwards.",
        ));
    }
    if from == to {
        return if op.get("expected").is_none_or(|expected| {
            expected
                .get("content")
                .and_then(Value::as_array)
                .is_some_and(Vec::is_empty)
        }) {
            Ok((paragraph.clone(), vec![]))
        } else {
            Err(Failure::refused(
                op,
                "stale",
                "The range to delete is empty.",
            ))
        };
    }
    let mut cursor = 0;
    for node in content {
        let end = cursor + inline::list_width(std::slice::from_ref(node));
        if node["type"] == "insertion"
            && ((cursor < from && from < end) || (cursor < to && to < end))
        {
            return Err(Failure::unsupported(op, "partialTrackedWrapperCuts"));
        }
        cursor = end;
    }
    let regions = inline::partition(content, &[from, to]);
    let removed = json!({"content":regions[1],
        "openStart":inline::spanning(content,&[from,to],0,1),
        "openEnd":inline::spanning(content,&[from,to],1,2)});
    if op
        .get("expected")
        .is_some_and(|expected| *expected != removed)
    {
        return Err(Failure::refused(
            op,
            "stale",
            format!(
                "The range of {} holds other content than expected.",
                op["from"]["blockId"].as_str().unwrap_or("")
            ),
        ));
    }
    let across = inline::spanning(content, &[from, to], 0, 2);
    let mut remaining = inline::merge(&regions[0], &regions[2], across, false)
        .expect("two pieces of the same plain record merge");
    if join > 0 {
        let depth = inline::spanning(&remaining, &[from], 0, 1) + join;
        let halves = inline::partition(&remaining, &[from]);
        remaining = inline::merge(&halves[0], &halves[1], depth, false).ok_or_else(|| {
            Failure::refused(
                op,
                "structureMismatch",
                format!("The records meeting at {from} cannot be merged."),
            )
        })?;
    }
    if remaining.iter().any(|node| node["type"] == "insertion") {
        // The inverse inserts its slice into the result. That supported
        // insertion boundary currently requires a plain source paragraph.
        return Err(Failure::unsupported(
            op,
            "trackedDeletionInverseWithRemainingWrappers",
        ));
    }
    let inverse = json!({"type":"insertContent","at":position(&op["from"],paragraph,from),
        "slice":removed});
    Ok((with_content(paragraph, remaining), vec![inverse]))
}

/// Reinsert complete wrappers returned by a supported direct deletion. The
/// source paragraph remains plain; open slice ends can continue plain runs
/// and text records, while wrapper ends must remain closed.
fn insert_wrapper_slice(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    let content = items(paragraph, op)?;
    inline::validate_plain(content, op, false)?;
    let slice = op.get("slice").and_then(Value::as_object).ok_or_else(|| {
        Failure::refused(op, "invalidOperation", "An insertion must name its slice.")
    })?;
    let incoming = slice
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOperation",
                "A slice must contain a content array.",
            )
        })?;
    if incoming.is_empty() {
        return Err(Failure::refused(
            op,
            "emptyContent",
            "The slice holds nothing.",
        ));
    }
    coordinate_view(incoming, op, true)?;
    let depth = |key: &str| -> Result<usize, Failure> {
        let value = slice
            .get(key)
            .and_then(Value::as_u64)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| Failure::refused(op, "structureMismatch", "Open depths are counts."))?;
        if value > 2 {
            return Err(Failure::unsupported(op, "trackedSliceOpenDepth"));
        }
        Ok(value)
    };
    let open_start = depth("openStart")?;
    let open_end = depth("openEnd")?;
    let from = inline::coordinate(content, &op["at"], op)?;
    let to = from + inline::list_width(incoming);
    let halves = inline::partition(content, &[from]);
    let changed_content = inline::merge(&halves[0], incoming, open_start, false)
        .and_then(|start| inline::merge(&start, &halves[1], open_end, true))
        .ok_or_else(|| {
            Failure::refused(
                op,
                "structureMismatch",
                format!(
                    "The slice's open ends do not fit {from} in {}.",
                    op["at"]["blockId"].as_str().unwrap_or("")
                ),
            )
        })?;
    // The whole wrappers came from the slice, so neither range endpoint cuts
    // one. The existing plain-run delete helpers preserve all open ends.
    let (restored, removed) = inline::delete_between(&changed_content, from, to);
    let cut = inline::spanning(content, &[from], 0, 1);
    let across = inline::spanning(&restored, &[from], 0, 1);
    if across > cut {
        return Err(Failure::refused(
            op,
            "structureMismatch",
            "The insertion would merge records that were separate.",
        ));
    }
    let mut inverse = json!({"type":"deleteRange","from":position(&op["at"],paragraph,from),
        "to":position(&op["at"],paragraph,to),"expected":removed});
    if cut > across {
        inverse["join"] = json!(cut - across);
    }
    Ok((with_content(paragraph, changed_content), vec![inverse]))
}

pub fn edit(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    if op
        .get("undefinedFields")
        .and_then(Value::as_array)
        .is_some_and(|paths| !paths.is_empty())
    {
        return Err(Failure::unsupported(op, "inlineOwnUndefinedFields"));
    }
    if op.get("revision").is_none() {
        return match op.get("type").and_then(Value::as_str) {
            Some("deleteRange") => delete_whole_insertions(paragraph, op),
            Some("insertContent") => insert_wrapper_slice(paragraph, op),
            _ => Err(Failure::unsupported(op, "trackedInlineOperationKind")),
        };
    }
    let content = items(paragraph, op)?;
    inline::validate_plain(content, op, false)?;
    let mut direct = op.clone();
    direct
        .as_object_mut()
        .expect("operation object")
        .remove("revision");
    match op.get("type").and_then(Value::as_str) {
        Some("insertText") => {
            let (inserted, direct_inverse) = inline::edit(paragraph, &direct)?;
            let from = inline::coordinate(content, &op["at"], op)?;
            let text = op["text"].as_str().expect("direct edit validated text");
            let to = from + text.encode_utf16().count();
            let inserted_content = items(&inserted, op)?;
            let direct_join = direct_inverse[0]
                .get("join")
                .and_then(Value::as_u64)
                .unwrap_or(0) as usize;
            let resolution_join =
                direct_join + inline::spanning(inserted_content, &[from, to], 0, 2);
            let wrapped = wrap_plain(
                inserted_content,
                from,
                to,
                "insertion",
                stamp_info(op)?,
                resolution_join,
            );
            let regions = inline::partition(&wrapped, &[from, to]);
            let removed = json!({"content":regions[1],"openStart":0,"openEnd":0});
            let mut inverse = json!({"type":"deleteRange", "from":position(&op["at"],paragraph,from),
                "to":position(&op["at"],paragraph,to),"expected":removed});
            // Wrapping separates the source records, so no record spans the
            // restored gap until this explicit join restores the original cut.
            let join = inline::spanning(content, &[from], 0, 1);
            if join > 0 {
                inverse["join"] = json!(join);
            }
            Ok((with_content(paragraph, wrapped), vec![inverse]))
        }
        Some("deleteRange") => {
            if op
                .get("join")
                .is_some_and(|join| join.as_u64().is_some_and(|join| join > 0))
            {
                return Err(Failure::refused(
                    op,
                    "structureMismatch",
                    "A tracked deletion joins nothing.",
                ));
            }
            let (_, direct_inverse) = inline::edit(paragraph, &direct)?;
            if direct_inverse.is_empty() {
                return Ok((paragraph.clone(), vec![]));
            }
            let from = inline::coordinate(content, &op["from"], op)?;
            let to = inline::coordinate(content, &op["to"], op)?;
            let wrapped = wrap_plain(
                content,
                from,
                to,
                "deletion",
                stamp_info(op)?,
                inline::spanning(content, &[from, to], 0, 2),
            );
            let changed = with_content(paragraph, wrapped);
            let inverse = restoring(paragraph, &changed, &op["from"]);
            Ok((changed, vec![inverse]))
        }
        _ => Err(Failure::unsupported(op, "trackedInlineOperationKind")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph() -> Value {
        json!({"type":"paragraph","paraId":"2F48B967","content":[
            {"type":"run","preservedAttributes":[{"name":"rsidR","value":"00E10F27"}],
                "content":[{"type":"text","text":"ž😀a","xmlSpace":"preserve"},
                    {"type":"text","text":"bc","extra":"authored"}]},
            {"type":"run","formatting":{"bold":true},"content":[{"type":"text","text":"DE"}]}]})
    }

    fn at(offset: usize) -> Value {
        json!({"story":"main","blockId":"2F48B967","offset":offset})
    }

    fn revision() -> Value {
        json!({"id":2147483001,"author":"Test Author","date":"2026-10-05T09:00:00Z","initials":"TA"})
    }

    fn scalar_boundaries(paragraph: &Value) -> Vec<usize> {
        let mut boundaries = vec![0];
        let mut offset = 0;
        fn visit(nodes: &[Value], offset: &mut usize, boundaries: &mut Vec<usize>) {
            for node in nodes {
                if node["type"] != "text" {
                    visit(node["content"].as_array().unwrap(), offset, boundaries);
                    continue;
                }
                for scalar in node["text"].as_str().unwrap().chars() {
                    *offset += scalar.len_utf16();
                    boundaries.push(*offset);
                }
            }
        }
        visit(
            paragraph["content"].as_array().unwrap(),
            &mut offset,
            &mut boundaries,
        );
        boundaries
    }

    #[test]
    fn all_scalar_insertions_have_exact_deletion_closure() {
        let original = paragraph();
        for offset in scalar_boundaries(&original) {
            for props in [json!("inherit"), json!({}), json!({"italic":true})] {
                let op = json!({"type":"insertText","at":at(offset),"text":"ř𐐀",
                    "runProps":props,"revision":revision()});
                let (changed, inverse) = edit(&original, &op).unwrap();
                let (restored, redo) = edit(&changed, &inverse[0]).unwrap();
                assert_eq!(restored, original, "offset {offset}, props {props}");
                let (reinserted, repeated_inverse) = edit(&restored, &redo[0]).unwrap();
                assert_eq!(reinserted, changed, "redo offset {offset}, props {props}");
                let (restored_again, _) = edit(&reinserted, &repeated_inverse[0]).unwrap();
                assert_eq!(restored_again, original);
                let wrapper = changed["content"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|node| node["type"] == "insertion")
                    .unwrap();
                assert_eq!(wrapper["info"], revision());
                assert_eq!(inverse[0]["expected"]["content"], json!([wrapper]));
                assert_eq!(inverse[0]["expected"]["openStart"], 0);
                assert_eq!(inverse[0]["expected"]["openEnd"], 0);
            }
        }
    }

    #[test]
    fn all_ranges_covering_a_whole_insertion_have_exact_slice_closure() {
        let original = paragraph();
        for offset in scalar_boundaries(&original) {
            let op = json!({"type":"insertText","at":at(offset),"text":"ř𐐀",
                "runProps":"inherit","revision":revision()});
            let (changed, _) = edit(&original, &op).unwrap();
            let boundaries = scalar_boundaries(&changed);
            for from in boundaries.iter().copied().filter(|from| *from <= offset) {
                for to in boundaries.iter().copied().filter(|to| *to >= offset + 3) {
                    let deletion = json!({"type":"deleteRange","from":at(from),"to":at(to)});
                    let (removed, inverse) = edit(&changed, &deletion).unwrap();
                    let (restored, _) = edit(&removed, &inverse[0]).unwrap();
                    assert_eq!(restored, changed, "insertion {offset}, range {from}..{to}");
                }
            }
        }
    }

    #[test]
    fn fresh_wrapper_info_uses_only_fields_written_by_ts_stamp_info() {
        let mut stamp = revision();
        stamp["utcDate"] = json!("2026-10-05T09:00:00Z");
        stamp["preservedAttributes"] = json!([{"name":"custom","value":"retained-on-source-only"}]);
        stamp.as_object_mut().unwrap().remove("initials");
        let op = json!({"type":"insertText","at":at(0),"text":"x",
            "runProps":"inherit","revision":stamp});
        let (changed, _) = edit(&paragraph(), &op).unwrap();
        let wrapper = changed["content"]
            .as_array()
            .unwrap()
            .iter()
            .find(|node| node["type"] == "insertion")
            .unwrap();
        assert_eq!(
            wrapper["info"],
            json!({"id":2147483001,"author":"Test Author",
            "date":"2026-10-05T09:00:00Z"})
        );
    }

    #[test]
    fn selective_two_wrapper_deletion_is_unsupported_until_its_inverse_can_replay() {
        let original = paragraph();
        let mut second_info = revision();
        second_info["id"] = json!(2147483002);
        let slice = json!({"content":[
            {"type":"insertion","info":revision(),
                "resolutionJoins":{"before":0,"after":0,"remove":0},
                "content":[{"type":"run","content":[{"type":"text","text":"x"}]}]},
            {"type":"insertion","info":second_info,
                "resolutionJoins":{"before":0,"after":0,"remove":0},
                "content":[{"type":"run","content":[{"type":"text","text":"y"}]}]}],
            "openStart":0,"openEnd":0});
        let insertion = json!({"type":"insertContent","at":at(0),"slice":slice});
        let (changed, _) = edit(&original, &insertion).unwrap();
        let before = changed.clone();
        let selective = json!({"type":"deleteRange","from":at(0),"to":at(1)});
        assert!(matches!(edit(&changed,&selective),
            Err(Failure::Unsupported {dimension,..})
            if dimension == "trackedDeletionInverseWithRemainingWrappers"));
        assert_eq!(changed, before);
        let all = json!({"type":"deleteRange","from":at(0),"to":at(2)});
        let (removed, inverse) = edit(&changed, &all).unwrap();
        assert_eq!(removed, original);
        let (restored, _) = edit(&removed, &inverse[0]).unwrap();
        assert_eq!(restored, changed);
    }

    #[test]
    fn all_scalar_deletions_restore_full_authored_content() {
        let original = paragraph();
        let boundaries = scalar_boundaries(&original);
        for from in boundaries.iter().copied() {
            for to in boundaries.iter().copied().filter(|to| *to > from) {
                let op =
                    json!({"type":"deleteRange","from":at(from),"to":at(to),"revision":revision()});
                let (changed, inverse) = edit(&original, &op).unwrap();
                assert_eq!(
                    inverse,
                    vec![json!({"type":"replaceInline","story":"main",
                    "blockId":"2F48B967","expected":changed["content"],"content":original["content"]})]
                );
                let wrapper = changed["content"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|node| node["type"] == "deletion")
                    .unwrap();
                assert_eq!(wrapper["info"], revision());
                assert_eq!(
                    wrapper["resolutionJoins"]["remove"],
                    inline::spanning(original["content"].as_array().unwrap(), &[from, to], 0, 2)
                );
                assert!(wrapper["resolutionJoins"].get("retainedAfter").is_none());
            }
        }
    }

    #[test]
    fn rejects_scalar_interiors_joined_deletions_and_partial_wrapper_cuts() {
        let original = paragraph();
        let op = json!({"type":"insertText","at":at(2),"text":"x","runProps":"inherit","revision":revision()});
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "splitsSurrogatePair")
        );
        let op =
            json!({"type":"deleteRange","from":at(0),"to":at(1),"join":1,"revision":revision()});
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "structureMismatch")
        );
        let op = json!({"type":"insertText","at":at(0),"text":"ab","runProps":"inherit","revision":revision()});
        let (changed, _) = edit(&original, &op).unwrap();
        let cut = json!({"type":"deleteRange","from":at(0),"to":at(1)});
        assert!(
            matches!(edit(&changed,&cut),Err(Failure::Unsupported {dimension,..}) if dimension == "partialTrackedWrapperCuts")
        );
    }
}
