//! Direct plain-run property patches and their compact TS restoring spans.
//! Changed cut runs remain cut. Unchanged cut pieces merge back immediately;
//! the inverse joins changed pieces after restoring their previous properties.

use crate::inline::{coordinate, list_width, merge, partition, spanning, validate_plain};
use crate::refusal::Failure;
use serde_json::{Map, Value, json};

#[derive(Clone)]
enum RunState {
    Unchanged,
    Changed { previous: Option<Value> },
}

#[derive(Clone)]
struct Piece {
    value: Value,
    state: RunState,
}

struct RestoreSpan {
    from: usize,
    to: usize,
    patch: Value,
    when_empty: &'static str,
    /// TS groups by JSON spelling, including the key order of nested values.
    key: String,
}

fn count(value: Option<&Value>, op: &Value) -> Result<usize, Failure> {
    match value {
        None => Ok(0),
        Some(value) => value
            .as_u64()
            .and_then(|number| usize::try_from(number).ok())
            .ok_or_else(|| Failure::refused(op, "structureMismatch", "A join depth is a count.")),
    }
}

fn property_set(value: Option<&Value>, op: &Value) -> Result<Option<Map<String, Value>>, Failure> {
    value
        .map(|value| {
            value
                .as_object()
                .cloned()
                .ok_or_else(|| Failure::unsupported(op, "runFormattingShape"))
        })
        .transpose()
}

fn states_values(formatting: Option<&Value>, expected: &Map<String, Value>) -> bool {
    expected.iter().all(|(key, value)| {
        let actual = formatting.and_then(|formatting| formatting.get(key));
        if value.is_null() {
            actual.is_none()
        } else {
            actual == Some(value)
        }
    })
}

fn patched_set(
    base: Option<&Value>,
    patch: &Map<String, Value>,
    keep_empty: bool,
    op: &Value,
) -> Result<Option<Value>, Failure> {
    let mut next = property_set(base, op)?.unwrap_or_default();
    for (key, value) in patch {
        if value.is_null() {
            next.remove(key);
        } else {
            next.insert(key.clone(), value.clone());
        }
    }
    Ok((!next.is_empty() || keep_empty).then_some(Value::Object(next)))
}

fn same_formatting(left: Option<&Value>, right: Option<&Value>) -> bool {
    let empty = json!({});
    left.unwrap_or(&empty) == right.unwrap_or(&empty)
}

fn prior_values(previous: Option<&Value>, patch: &Map<String, Value>) -> Value {
    let restore = patch
        .keys()
        .map(|key| {
            (
                key.clone(),
                previous
                    .and_then(|previous| previous.get(key))
                    .cloned()
                    .unwrap_or(Value::Null),
            )
        })
        .collect::<Map<_, _>>();
    Value::Object(restore)
}

fn with_formatting(run: &Value, formatting: Option<Value>) -> Value {
    let mut next = run.clone();
    let fields = next.as_object_mut().expect("validated run");
    fields.remove("formatting");
    if let Some(formatting) = formatting {
        fields.insert("formatting".into(), formatting);
    }
    next
}

fn unchanged(items: Vec<Value>) -> Vec<Piece> {
    items
        .into_iter()
        .map(|value| Piece {
            value,
            state: RunState::Unchanged,
        })
        .collect()
}

fn values(pieces: &[Piece]) -> Vec<Value> {
    pieces.iter().map(|piece| piece.value.clone()).collect()
}

/// `asFarAsAlike` on the supported two-level tree: a changed property set
/// stops at the run; unchanged pieces of an authored run merge completely.
fn merge_alike(left: &[Piece], right: &[Piece], depth: usize) -> Vec<Piece> {
    if depth == 0 {
        return left.iter().chain(right).cloned().collect();
    }
    let Some(merged) = merge(&values(left), &values(right), depth, false) else {
        return left.iter().chain(right).cloned().collect();
    };
    let last = left.last().expect("successful merge has left edge");
    let first = right.first().expect("successful merge has right edge");
    assert!(
        matches!(last.state, RunState::Unchanged) && matches!(first.state, RunState::Unchanged),
        "a changed run's formatting differs from its untouched cut piece"
    );
    let mut pieces = left[..left.len() - 1].to_vec();
    pieces.push(Piece {
        value: merged[left.len() - 1].clone(),
        state: RunState::Unchanged,
    });
    pieces.extend_from_slice(&right[1..]);
    pieces
}

fn patch_between(
    items: &[Value],
    from: usize,
    to: usize,
    patch: &Map<String, Value>,
    keep_empty: bool,
    op: &Value,
) -> Result<Option<(Vec<Value>, Vec<RestoreSpan>)>, Failure> {
    let regions = partition(items, &[from, to]);
    let before = unchanged(regions[0].clone());
    let after = unchanged(regions[2].clone());
    let mut changed = false;
    let mut middle = Vec::new();
    for run in &regions[1] {
        let previous = run.get("formatting");
        let formatting = patched_set(previous, patch, keep_empty, op)?;
        if same_formatting(formatting.as_ref(), previous) {
            middle.push(Piece {
                value: run.clone(),
                state: RunState::Unchanged,
            });
            continue;
        }
        changed = true;
        middle.push(Piece {
            value: with_formatting(run, formatting),
            state: RunState::Changed {
                previous: previous.cloned(),
            },
        });
    }
    if !changed {
        return Ok(None);
    }
    let head = merge_alike(&before, &middle, spanning(items, &[from, to], 0, 1));
    let whole = merge_alike(&head, &after, spanning(items, &[from, to], 1, 2));
    let mut restoring: Vec<RestoreSpan> = Vec::new();
    let mut stretch_open = false;
    let mut offset = 0;
    for piece in &whole {
        let start = offset;
        offset += list_width(std::slice::from_ref(&piece.value));
        let RunState::Changed { previous } = &piece.state else {
            stretch_open = false;
            continue;
        };
        let restore = prior_values(previous.as_ref(), patch);
        let spelling = if previous
            .as_ref()
            .and_then(Value::as_object)
            .is_some_and(Map::is_empty)
        {
            "keep"
        } else {
            "omit"
        };
        let key = serde_json::to_string(&json!([restore, spelling]))
            .expect("property snapshot serializes");
        if stretch_open && restoring.last().is_some_and(|last| last.key == key) {
            restoring.last_mut().expect("checked restoring span").to = offset;
        } else {
            restoring.push(RestoreSpan {
                from: start,
                to: offset,
                patch: restore,
                when_empty: spelling,
                key,
            });
        }
        stretch_open = true;
    }
    Ok(Some((values(&whole), restoring)))
}

fn position(source: &Value, paragraph: &Value, offset: usize) -> Value {
    json!({"story":source["story"],"blockId":paragraph["paraId"],"offset":offset,"zeroWidthBefore":0})
}

pub fn edit(paragraph: &Value, op: &Value) -> Result<(Value, Vec<Value>), Failure> {
    let join_start = count(op.get("joinStart"), op)?;
    let join_end = count(op.get("joinEnd"), op)?;
    if op.get("revision").is_some() {
        return Err(if join_start != 0 || join_end != 0 {
            Failure::refused(op, "structureMismatch", "A tracked patch joins nothing.")
        } else {
            Failure::unsupported(op, "trackedRunPropertyPatch")
        });
    }
    if op
        .get("undefinedFields")
        .and_then(Value::as_array)
        .is_some_and(|fields| !fields.is_empty())
    {
        return Err(Failure::unsupported(op, "runPatchOwnUndefinedFields"));
    }
    let items = paragraph
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "paragraphContentShape"))?;
    validate_plain(items, op, false)?;
    let from = coordinate(items, &op["from"], op)?;
    let to = coordinate(items, &op["to"], op)?;
    if from > to {
        return Err(Failure::refused(
            op,
            "invalidOffset",
            "The range runs backwards.",
        ));
    }
    if from == to {
        return Ok((paragraph.clone(), vec![]));
    }
    let patch = op.get("patch").and_then(Value::as_object).ok_or_else(|| {
        Failure::refused(
            op,
            "invalidOperation",
            "A run-property operation must name its patch.",
        )
    })?;
    if let Some(expected) = op.get("expected") {
        let expected = expected.as_object().ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOperation",
                "Expected run properties must be an object.",
            )
        })?;
        let selected = partition(items, &[from, to]);
        if !selected[1]
            .iter()
            .all(|run| states_values(run.get("formatting"), expected))
        {
            return Err(Failure::refused(
                op,
                "stale",
                "The runs state other values than expected.",
            ));
        }
    }
    let keep_empty = op.get("whenEmpty").and_then(Value::as_str) == Some("keep");
    let patched = patch_between(items, from, to, patch, keep_empty, op)?;
    let (mut content, restoring) = patched.unwrap_or_else(|| (items.clone(), vec![]));
    for (gap, join) in [(from, join_start), (to, join_end)] {
        if join == 0 {
            continue;
        }
        let across = spanning(&content, &[gap], 0, 1);
        let depth = across.checked_add(join).ok_or_else(|| {
            Failure::refused(
                op,
                "structureMismatch",
                "The records meeting at the gap cannot be merged.",
            )
        })?;
        let halves = partition(&content, &[gap]);
        content = merge(&halves[0], &halves[1], depth, false).ok_or_else(|| {
            Failure::refused(
                op,
                "structureMismatch",
                "The records meeting at the gap cannot be merged.",
            )
        })?;
    }
    if content == *items {
        return Ok((paragraph.clone(), vec![]));
    }
    let mut inverse = restoring.into_iter().map(|span|json!({"type":"setRunProps",
        "from":position(&op["from"],paragraph,span.from),"to":position(&op["from"],paragraph,span.to),
        "patch":span.patch,"whenEmpty":span.when_empty,"expected":op["patch"]})).collect::<Vec<_>>();
    if !inverse.is_empty() {
        let cut_from =
            spanning(items, &[from], 0, 1).saturating_sub(spanning(&content, &[from], 0, 1));
        let cut_to = spanning(items, &[to], 0, 1).saturating_sub(spanning(&content, &[to], 0, 1));
        if cut_from > 0 {
            inverse[0]["joinStart"] = json!(cut_from);
        }
        if cut_to > 0 {
            inverse.last_mut().expect("nonempty inverse")["joinEnd"] = json!(cut_to);
        }
    }
    let mut result = paragraph.clone();
    result["content"] = Value::Array(content);
    Ok((result, inverse))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph() -> Value {
        json!({"type":"paragraph","paraId":"00000001","content":[
            {"type":"run","preservedAttributes":[{"name":"rsidR","value":"abc"}],"content":[{"type":"text","text":"a😀b","opaque":"first"}]},
            {"type":"run","formatting":{},"content":[{"type":"text","text":"cd","opaque":"second"}]},
            {"type":"run","formatting":{"bold":false,"color":{"rgb":"FF0000"}},"content":[{"type":"text","text":"ef"}]},
            {"type":"run","formatting":{"bold":true},"content":[{"type":"text","text":"gh"}]}]})
    }

    fn position(offset: usize) -> Value {
        json!({"story":"main","blockId":"00000001","offset":offset})
    }

    fn restore(mut paragraph: Value, inverse: &[Value]) -> Value {
        for op in inverse {
            paragraph = edit(&paragraph, op).unwrap().0;
        }
        paragraph
    }

    #[test]
    fn every_scalar_range_and_patch_roundtrips_field_presence_and_authored_boundaries() {
        let original = paragraph();
        let boundaries = [0, 1, 3, 4, 5, 6, 7, 8, 9, 10];
        for from in boundaries {
            for to in boundaries.into_iter().filter(|to| *to > from) {
                for patch in [
                    json!({"bold":true}),
                    json!({"bold":false}),
                    json!({"bold":null}),
                    json!({"color":{"rgb":"00FF00"},"italic":true}),
                ] {
                    for when_empty in ["omit", "keep"] {
                        let op = json!({"type":"setRunProps","from":position(from),"to":position(to),"patch":patch,"whenEmpty":when_empty});
                        let (changed, inverse) = edit(&original, &op).unwrap();
                        assert_eq!(
                            restore(changed, &inverse),
                            original,
                            "range {from}..{to}, patch {patch}, {when_empty}"
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn grouped_restore_spans_stop_at_unchanged_runs_and_distinguish_empty_spelling() {
        let original = paragraph();
        let op = json!({"type":"setRunProps","from":position(0),"to":position(10),"patch":{"bold":true}});
        let (_, inverse) = edit(&original, &op).unwrap();
        assert_eq!(inverse.len(), 3);
        assert_eq!(inverse[0]["whenEmpty"], "omit");
        assert_eq!(inverse[1]["whenEmpty"], "keep");
        assert_eq!(inverse[2]["patch"]["bold"], false);
        assert_eq!(inverse[2]["to"]["offset"], 8);
    }

    #[test]
    fn stale_expected_properties_and_empty_ranges_follow_ts_semantics() {
        let original = paragraph();
        let op = json!({"type":"setRunProps","from":position(0),"to":position(4),"patch":{"italic":true},"expected":{"bold":false}});
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "stale")
        );
        let op = json!({"type":"setRunProps","from":position(1),"to":position(1),"patch":{"italic":true},"expected":{"bold":false}});
        assert_eq!(edit(&original, &op).unwrap(), (original, vec![]));
    }

    #[test]
    fn restoring_operations_merge_both_cut_edges_at_the_declared_depth() {
        let original = paragraph();
        let op = json!({"type":"setRunProps","from":position(1),"to":position(3),"patch":{"italic":true}});
        let (changed, inverse) = edit(&original, &op).unwrap();
        assert_eq!(inverse[0]["joinStart"], 2);
        assert_eq!(inverse[0]["joinEnd"], 2);
        assert_eq!(restore(changed, &inverse), original);
    }
}
