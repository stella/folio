//! Direct paragraph split/join, with the same ordered inverse operations as TS.
//! Inline cuts reuse the plain run/text partition kernel; review metadata is
//! moved with the paragraph mark rather than copied into both halves.

use crate::apply::{Applied, Path, Step, Touched, at_path, at_path_mut, locate};
use crate::inline::{coordinate, list_width, merge, partition, spanning, validate_plain};
use crate::refusal::Failure;
use serde_json::{Map, Value, json};
use std::collections::{BTreeMap, BTreeSet};

fn content(paragraph: &Value) -> &[Value] {
    paragraph["content"]
        .as_array()
        .expect("validated paragraph content")
}

fn validate_paragraph(paragraph: &Value, op: &Value) -> Result<(), Failure> {
    let content = paragraph
        .get("content")
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "paragraphContentShape"))?;
    validate_plain(content, op, false)
}

fn with_content(paragraph: &Value, content: Vec<Value>) -> Value {
    let mut next = paragraph.clone();
    next["content"] = Value::Array(content);
    next
}

fn locator(op: &Value, story: &Value, block_id: &Value) -> Value {
    json!({"type":op["type"],"story":story,"blockId":block_id})
}

fn list_and_index(path: &[Step], op: &Value) -> Result<(Path, usize), Failure> {
    let Some(Step::Index(index)) = path.last() else {
        return Err(Failure::unsupported(op, "paragraphBlockListShape"));
    };
    Ok((path[..path.len() - 1].to_vec(), *index))
}

fn valid_id(id: &str) -> bool {
    id.len() == 8
        && id.bytes().all(|byte| byte.is_ascii_hexdigit())
        && u32::from_str_radix(id, 16).is_ok_and(|number| number > 0 && number < 0x80000000)
}

fn all_ids(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::Object(fields) => {
            if fields.get("type").and_then(Value::as_str) == Some("paragraph") {
                if let Some(id) = fields.get("paraId").and_then(Value::as_str) {
                    out.push(id.to_owned());
                }
            }
            for child in fields.values() {
                all_ids(child, out);
            }
        }
        Value::Array(values) => {
            for value in values {
                all_ids(value, out);
            }
        }
        _ => {}
    }
}

fn package_ids(document: &Value) -> Vec<String> {
    let mut package = document["package"].clone();
    if let Some(body) = package.get_mut("document").and_then(Value::as_object_mut) {
        body.remove("sections");
    }
    let mut ids = Vec::new();
    all_ids(&package, &mut ids);
    ids
}

fn split_fields(paragraph: &Value) -> Value {
    let mut fields = paragraph.as_object().expect("validated paragraph").clone();
    for key in ["type", "paraId", "content", "sectionProperties", "pPrMark"] {
        fields.remove(key);
    }
    Value::Object(fields)
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

fn with_review(paragraph: &Value, review: &Value) -> Value {
    let mut next = paragraph.clone();
    let fields = next.as_object_mut().expect("validated paragraph");
    for key in ["formatting", "propertyChanges", "pPrMark"] {
        fields.remove(key);
        if let Some(value) = review.get(key) {
            fields.insert(key.into(), value.clone());
        }
    }
    next
}

fn joined_formatting(first: &Value, second: &Value) -> Option<Value> {
    if list_width(content(first)) == 0 {
        return second.get("formatting").cloned();
    }
    let properties = first.get("formatting").and_then(Value::as_object);
    let marks = second.get("formatting").and_then(Value::as_object);
    let is_mark = |key: &str| matches!(key, "runProperties" | "runInWithNext");
    let mut fields = properties
        .map(|properties| {
            properties
                .iter()
                .filter(|(key, _)| !is_mark(key))
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect::<Map<_, _>>()
        })
        .unwrap_or_default();
    let mut has_marks = false;
    if let Some(marks) = marks {
        for (key, value) in marks.iter().filter(|(key, _)| is_mark(key)) {
            has_marks = true;
            fields.insert(key.clone(), value.clone());
        }
    }
    (properties.is_some() || has_marks).then_some(Value::Object(fields))
}

fn joined_review(first: &Value, second: &Value) -> Value {
    let mut review = Map::new();
    if let Some(formatting) = joined_formatting(first, second) {
        review.insert("formatting".into(), formatting);
    }
    for key in ["propertyChanges", "pPrMark"] {
        if let Some(value) = second.get(key) {
            review.insert(key.into(), value.clone());
        }
    }
    Value::Object(review)
}

fn review_setting(story: &Value, id: &Value, from: Value, to: Value) -> Vec<Value> {
    if from == to {
        return vec![];
    }
    vec![
        json!({"type":"setParagraphReview","story":story,"blockId":id,"expected":from,"review":to}),
    ]
}

fn content_restoring(story: &Value, id: &Value, from: &[Value], to: &[Value]) -> Value {
    json!({"type":"replaceInline","story":story,"blockId":id,"expected":from,"content":to})
}

fn equal_retired(actual: &Value, expected: &Value) -> bool {
    let without_derived = |value: &Value| {
        let mut fields = value.as_object().cloned().unwrap_or_default();
        fields.remove("listRendering");
        fields.remove("renderedPageBreakBefore");
        fields
    };
    without_derived(actual) == without_derived(expected)
}

fn half(value: Option<&Value>, default: &'static str, op: &Value) -> Result<&'static str, Failure> {
    match value.and_then(Value::as_str) {
        Some("first") => Ok("first"),
        Some("second") => Ok("second"),
        None if value.is_none() => Ok(default),
        _ => Err(Failure::refused(
            op,
            "structureMismatch",
            "A half is first or second.",
        )),
    }
}

fn count(value: Option<&Value>, op: &Value) -> Result<usize, Failure> {
    match value {
        None => Ok(0),
        Some(value) => value
            .as_u64()
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| Failure::refused(op, "structureMismatch", "A join depth is a count.")),
    }
}

fn field<'a>(value: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    value.and_then(|value| value.get(key))
}

fn slots_equal(left: Option<&Value>, right: Option<&Value>, keys: &[&str]) -> bool {
    keys.iter().all(|key| field(left, key) == field(right, key))
}

fn optional_object_equal(left: Option<&Value>, right: Option<&Value>, keys: &[&str]) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => slots_equal(Some(left), Some(right), keys),
        _ => false,
    }
}

fn colors_equal(left: Option<&Value>, right: Option<&Value>) -> bool {
    optional_object_equal(
        left,
        right,
        &["rgb", "auto", "themeColor", "themeTint", "themeShade"],
    )
}

/// Exact TS runMerge formatting comparison, including absent/empty distinctions.
fn formatting_equal(left: Option<&Value>, right: Option<&Value>) -> bool {
    let (left, right) = match (left, right) {
        (None, None) => return true,
        (Some(left), Some(right)) => (left, right),
        _ => return false,
    };
    if !slots_equal(
        Some(left),
        Some(right),
        &[
            "bold",
            "boldCs",
            "italic",
            "italicCs",
            "strike",
            "doubleStrike",
            "vertAlign",
            "smallCaps",
            "allCaps",
            "hidden",
            "noProof",
            "highlight",
            "fontSize",
            "fontSizeCs",
            "spacing",
            "position",
            "scale",
            "kerning",
            "effect",
            "emphasisMark",
            "emboss",
            "imprint",
            "outline",
            "shadow",
            "rtl",
            "cs",
            "styleId",
        ],
    ) {
        return false;
    }
    let underline = (left.get("underline"), right.get("underline"));
    if !optional_object_equal(underline.0, underline.1, &["style"])
        || !colors_equal(field(underline.0, "color"), field(underline.1, "color"))
        || !colors_equal(left.get("color"), right.get("color"))
    {
        return false;
    }
    let shading = (left.get("shading"), right.get("shading"));
    if !optional_object_equal(shading.0, shading.1, &["pattern"])
        || !colors_equal(field(shading.0, "color"), field(shading.1, "color"))
        || !colors_equal(field(shading.0, "fill"), field(shading.1, "fill"))
    {
        return false;
    }
    if !optional_object_equal(
        left.get("fontFamily"),
        right.get("fontFamily"),
        &[
            "ascii",
            "hAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hAnsiTheme",
            "eastAsiaTheme",
            "csTheme",
        ],
    ) || !slots_equal(
        left.get("language"),
        right.get("language"),
        &["val", "eastAsia", "bidi"],
    ) {
        return false;
    }
    let preserved = |value: &Value| {
        value
            .get("preserved")
            .and_then(|value| value.get("children"))
            .and_then(Value::as_array)
            .map(|children| {
                children
                    .iter()
                    .map(|child| (child.get("index").cloned(), child.get("xml").cloned()))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
    };
    preserved(left) == preserved(right)
}

fn attribute_keys(run: &Value) -> Vec<String> {
    let mut keys = run
        .get("preservedAttributes")
        .and_then(Value::as_array)
        .map(|attributes| {
            attributes
                .iter()
                .map(|attribute| {
                    format!(
                        "{}\0{}\0{}",
                        attribute
                            .get("namespace")
                            .and_then(Value::as_str)
                            .unwrap_or(""),
                        attribute.get("name").and_then(Value::as_str).unwrap_or(""),
                        attribute.get("value").and_then(Value::as_str).unwrap_or("")
                    )
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    keys.sort();
    keys
}

/// The parser's plain-run seam rule, deliberately different from exact open-end
/// merges: the boundary text record is reconstructed as `{type,text}`.
fn paragraph_seam(left: &[Value], right: &[Value]) -> Vec<Value> {
    let (Some(last), Some(first)) = (left.last(), right.first()) else {
        return left.iter().chain(right).cloned().collect();
    };
    if !formatting_equal(last.get("formatting"), first.get("formatting"))
        || attribute_keys(last) != attribute_keys(first)
    {
        return left.iter().chain(right).cloned().collect();
    }
    let left_children = content(last);
    let right_children = content(first);
    let last_text = left_children.last().expect("nonempty validated run");
    let first_text = right_children.first().expect("nonempty validated run");
    let mut children = left_children[..left_children.len() - 1].to_vec();
    children.push(json!({"type":"text","text":format!("{}{}",last_text["text"].as_str().expect("validated text"),first_text["text"].as_str().expect("validated text"))}));
    children.extend_from_slice(&right_children[1..]);
    let merged = with_content(last, children);
    let mut result = left[..left.len() - 1].to_vec();
    result.push(merged);
    result.extend_from_slice(&right[1..]);
    result
}

fn ensure_identity_subset(
    document: &Value,
    before: &[Value],
    after: &[Value],
    op: &Value,
) -> Result<(), Failure> {
    let before_keys = crate::identity::keys(&json!(before));
    let after_keys = crate::identity::keys(&json!(after));
    let mut outside = BTreeMap::<String, usize>::new();
    for key in crate::identity::package_keys(document) {
        *outside.entry(key).or_default() += 1;
    }
    for key in before_keys {
        if let Some(count) = outside.get_mut(&key) {
            *count = count.saturating_sub(1);
        }
    }
    let mut seen = BTreeSet::new();
    if after_keys
        .iter()
        .any(|key| outside.get(key).is_some_and(|count| *count > 0) || !seen.insert(key))
    {
        return Err(Failure::unsupported(op, "paragraphIdentityFreshening"));
    }
    Ok(())
}

fn commit(
    document: &Value,
    op: &Value,
    list: &[Step],
    start: usize,
    before: &[Value],
    after: Vec<Value>,
    inverse: Vec<Value>,
) -> Result<Applied, Failure> {
    ensure_identity_subset(document, before, &after, op)?;
    let before_ids = before
        .iter()
        .map(|p| {
            p["paraId"]
                .as_str()
                .expect("located paragraph id")
                .to_owned()
        })
        .collect::<Vec<_>>();
    let after_ids = after
        .iter()
        .map(|p| {
            p["paraId"]
                .as_str()
                .expect("constructed paragraph id")
                .to_owned()
        })
        .collect::<Vec<_>>();
    let old_keys = before_ids
        .iter()
        .map(|id| id.to_ascii_uppercase())
        .collect::<BTreeSet<_>>();
    let new_keys = after_ids
        .iter()
        .map(|id| id.to_ascii_uppercase())
        .collect::<BTreeSet<_>>();
    let touched = Touched {
        modified: before_ids
            .iter()
            .filter(|id| new_keys.contains(&id.to_ascii_uppercase()))
            .cloned()
            .collect(),
        inserted: after_ids
            .iter()
            .filter(|id| !old_keys.contains(&id.to_ascii_uppercase()))
            .cloned()
            .collect(),
        removed: before_ids
            .iter()
            .filter(|id| !new_keys.contains(&id.to_ascii_uppercase()))
            .cloned()
            .collect(),
    };
    let mut result = document.clone();
    let blocks = at_path_mut(&mut result, list)
        .and_then(Value::as_array_mut)
        .ok_or_else(|| Failure::unsupported(op, "paragraphBlockListShape"))?;
    blocks.splice(start..start + before.len(), after);
    Ok(Applied {
        document: result,
        inverse,
        touched,
        revisions: vec![],
    })
}

fn split(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let new_id = op
        .get("newBlockId")
        .and_then(Value::as_str)
        .filter(|id| valid_id(id))
        .ok_or_else(|| {
            Failure::refused(op, "invalidBlockId", "The new id is not a paragraph id.")
        })?;
    if package_ids(document)
        .iter()
        .any(|id| id.eq_ignore_ascii_case(new_id))
    {
        return Err(Failure::refused(
            op,
            "idCollision",
            "The new paragraph id is already used in the package.",
        ));
    }
    if op.get("newHalf").is_some() {
        half(op.get("newHalf"), "first", op)?;
    }
    let at = op.get("at").ok_or_else(|| {
        Failure::refused(op, "invalidOperation", "A split must name its position.")
    })?;
    let path = locate(document, &locator(op, &at["story"], &at["blockId"]))?;
    let paragraph = at_path(document, &path).expect("located paragraph");
    validate_paragraph(paragraph, op)?;
    let (list, index) = list_and_index(&path, op)?;
    let offset = coordinate(content(paragraph), at, op)?;
    let new_half = half(
        op.get("newHalf"),
        if offset == list_width(content(paragraph)) {
            "second"
        } else {
            "first"
        },
        op,
    )?;
    if new_half == "second"
        && op
            .get("newParagraph")
            .and_then(|p| p.get("propertyChanges"))
            .and_then(Value::as_array)
            .is_some_and(|changes| !changes.is_empty())
    {
        return Err(Failure::refused(
            op,
            "structureMismatch",
            "The second half takes the paragraph's property changes with its mark.",
        ));
    }
    let fields = match op.get("newParagraph") {
        Some(value) => value.as_object().cloned().ok_or_else(|| {
            Failure::refused(
                op,
                "invalidOperation",
                "New paragraph fields must be an object.",
            )
        })?,
        None => {
            let mut fields = Map::new();
            if let Some(formatting) = paragraph.get("formatting") {
                fields.insert("formatting".into(), formatting.clone());
            }
            fields
        }
    };
    let cut = partition(content(paragraph), &[offset]);
    let through = spanning(content(paragraph), &[offset], 0, 1);
    let mut made = Value::Object(fields);
    made["type"] = json!("paragraph");
    made["paraId"] = json!(new_id);
    made["content"] = json!([]);
    let (mut first, second) = if new_half == "first" {
        (
            with_content(&made, cut[0].clone()),
            with_content(paragraph, cut[1].clone()),
        )
    } else {
        let mut first = with_content(paragraph, cut[0].clone());
        for key in ["sectionProperties", "pPrMark", "propertyChanges"] {
            first.as_object_mut().expect("paragraph").remove(key);
        }
        let mut second = with_content(&made, cut[1].clone());
        for key in ["propertyChanges", "sectionProperties", "pPrMark"] {
            if let Some(value) = paragraph.get(key) {
                second[key] = value.clone();
            }
        }
        (first, second)
    };
    if let Some(mark) = op.get("firstMark") {
        first["pPrMark"] = mark.clone();
    }
    if let Some(properties) = op.get("firstSectionProperties") {
        first["sectionProperties"] = properties.clone();
    }
    let (kept, made) = if new_half == "first" {
        (&second, &first)
    } else {
        (&first, &second)
    };
    let mut join = json!({"type":"joinBlocks","story":at["story"],"blockId":first["paraId"],"nextBlockId":second["paraId"],
        "depth":through,"survivor":if new_half == "first" {"second"} else {"first"},
        "expectedRetired":split_fields(made),"expectedSurvivor":review_fields(kept)});
    if first.get("sectionProperties").is_some() {
        join["sectionBoundary"] = json!("remove");
    }
    let inverse_content = if through == 0 {
        paragraph_seam(content(&first), content(&second))
    } else {
        content(paragraph).to_vec()
    };
    let mut inverse = vec![join];
    if inverse_content != content(paragraph) {
        inverse.push(content_restoring(
            &at["story"],
            &kept["paraId"],
            &inverse_content,
            content(paragraph),
        ));
    }
    inverse.extend(review_setting(
        &at["story"],
        &kept["paraId"],
        joined_review(&first, &second),
        review_fields(paragraph),
    ));
    commit(
        document,
        op,
        &list,
        index,
        &[paragraph.clone()],
        vec![first, second],
        inverse,
    )
}

fn join(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let path = locate(document, op)?;
    let next_path = locate(document, &locator(op, &op["story"], &op["nextBlockId"]))?;
    let (list, index) = list_and_index(&path, op)?;
    let (next_list, next_index) = list_and_index(&next_path, op)?;
    if list != next_list || next_index != index + 1 {
        return Err(Failure::refused(
            op,
            "notAdjacent",
            "The second paragraph does not directly follow the first.",
        ));
    }
    let first = at_path(document, &path).expect("located paragraph");
    let second = at_path(document, &next_path).expect("located paragraph");
    if first.get("sectionProperties").is_some()
        && op.get("sectionBoundary").and_then(Value::as_str) != Some("remove")
    {
        return Err(Failure::refused(
            op,
            "sectionBoundary",
            "The first paragraph ends a section.",
        ));
    }
    let survivor_side = half(op.get("survivor"), "second", op)?;
    let (survivor, retired) = if survivor_side == "second" {
        (second, first)
    } else {
        (first, second)
    };
    if op
        .get("expectedRetired")
        .is_some_and(|expected| !equal_retired(&split_fields(retired), expected))
        || op
            .get("expectedSurvivor")
            .is_some_and(|expected| review_fields(survivor) != *expected)
    {
        return Err(Failure::refused(
            op,
            "stale",
            "The paragraph fields differ from the join's expectation.",
        ));
    }
    let retired_id = retired["paraId"]
        .as_str()
        .filter(|id| valid_id(id))
        .ok_or_else(|| {
            Failure::refused(
                op,
                "invalidBlockId",
                "The retired paragraph id could not name it again.",
            )
        })?;
    validate_paragraph(first, op)?;
    validate_paragraph(second, op)?;
    let depth = count(op.get("depth"), op)?;
    let merged = if depth == 0 {
        paragraph_seam(content(first), content(second))
    } else {
        merge(content(first), content(second), depth, false).ok_or_else(|| {
            Failure::refused(
                op,
                "structureMismatch",
                "The records meeting between paragraphs cannot be merged.",
            )
        })?
    };
    let mut joined = with_review(
        &with_content(survivor, merged.clone()),
        &joined_review(first, second),
    );
    joined
        .as_object_mut()
        .expect("paragraph")
        .remove("sectionProperties");
    if let Some(properties) = second.get("sectionProperties") {
        joined["sectionProperties"] = properties.clone();
    }
    let length = list_width(content(first));
    let new_half = if survivor_side == "second" {
        "first"
    } else {
        "second"
    };
    let mut retired_fields = split_fields(retired);
    if new_half == "second" {
        retired_fields
            .as_object_mut()
            .expect("fields")
            .remove("propertyChanges");
    }
    let mut split = json!({"type":"splitBlock","at":{"story":op["story"],"blockId":survivor["paraId"],"offset":length,"zeroWidthBefore":0},
        "newBlockId":retired_id,"newHalf":new_half,"newParagraph":retired_fields});
    if let Some(mark) = first.get("pPrMark") {
        split["firstMark"] = mark.clone();
    }
    if let Some(properties) = first.get("sectionProperties") {
        split["firstSectionProperties"] = properties.clone();
    }
    let mut split_review = review_fields(&joined);
    if survivor_side == "first" {
        let fields = split_review.as_object_mut().expect("review fields");
        fields.remove("propertyChanges");
        fields.remove("pPrMark");
        if let Some(mark) = first.get("pPrMark") {
            fields.insert("pPrMark".into(), mark.clone());
        }
    }
    let mut inverse = vec![split];
    if depth == 0 && merged.len() < content(first).len() + content(second).len() {
        let restored = partition(&merged, &[length]);
        if restored[0] != content(first) {
            inverse.push(content_restoring(
                &op["story"],
                &first["paraId"],
                &restored[0],
                content(first),
            ));
        }
        if restored[1] != content(second) {
            inverse.push(content_restoring(
                &op["story"],
                &second["paraId"],
                &restored[1],
                content(second),
            ));
        }
    }
    inverse.extend(review_setting(
        &op["story"],
        &survivor["paraId"],
        split_review,
        review_fields(survivor),
    ));
    commit(
        document,
        op,
        &list,
        index,
        &[first.clone(), second.clone()],
        vec![joined],
        inverse,
    )
}

pub fn edit(document: &Value, op: &Value) -> Result<Applied, Failure> {
    if op.get("revision").is_some() {
        return Err(Failure::unsupported(op, "trackedParagraphSplitJoin"));
    }
    if op.get("sectionView").is_some() {
        return Err(Failure::unsupported(op, "explicitSectionViewRestoration"));
    }
    match op.get("type").and_then(Value::as_str) {
        Some("splitBlock") => split(document, op),
        Some("joinBlocks") => join(document, op),
        _ => Err(Failure::unsupported(op, "paragraphStructuralOperationKind")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph(id: &str, text: &str) -> Value {
        json!({"type":"paragraph","paraId":id,"textId":"authored","preservedAttributes":[{"name":"rsidR","value":"abc"}],
            "formatting":{"alignment":"center","runProperties":{"italic":true}},"content":if text.is_empty() {json!([])} else {
                json!([{"type":"run","content":[{"type":"text","text":text,"opaque":"leaf"}]}])}})
    }

    fn document(first: Value, second: Option<Value>) -> Value {
        let mut content = vec![first];
        if let Some(second) = second {
            content.push(second);
        }
        json!({"package":{"document":{"content":content}}})
    }

    fn apply_inverse(document: &Value, inverse: &[Value]) -> Value {
        crate::apply::apply_operations(document, inverse)
            .unwrap()
            .document
    }

    #[test]
    fn every_scalar_split_half_roundtrips_authored_fields_and_inverse_operations() {
        let original = document(paragraph("00000001", "a😀bc"), None);
        for offset in [0, 1, 3, 4, 5] {
            for half in ["first", "second"] {
                let op = json!({"type":"splitBlock","at":{"story":"main","blockId":"00000001","offset":offset},"newBlockId":"00000002","newHalf":half});
                let changed = edit(&original, &op).unwrap();
                assert_eq!(changed.inverse[0]["type"], "joinBlocks");
                assert_eq!(changed.touched.modified, ["00000001"]);
                assert_eq!(changed.touched.inserted, ["00000002"]);
                assert_eq!(
                    apply_inverse(&changed.document, &changed.inverse),
                    original,
                    "offset {offset}, half {half}"
                );
            }
        }
    }

    #[test]
    fn both_join_survivors_restore_seam_metadata_boundaries_and_review_fields() {
        for first_text in ["", "ab"] {
            for second_text in ["", "cd"] {
                let first = paragraph("00000001", first_text);
                let mut second = paragraph("00000002", second_text);
                second["formatting"] = json!({"alignment":"right","runProperties":{"bold":true}});
                let original = document(first, Some(second));
                for survivor in ["first", "second"] {
                    let op = json!({"type":"joinBlocks","story":"main","blockId":"00000001","nextBlockId":"00000002","survivor":survivor});
                    let changed = edit(&original, &op).unwrap();
                    assert_eq!(changed.inverse[0]["type"], "splitBlock");
                    assert_eq!(
                        apply_inverse(&changed.document, &changed.inverse),
                        original,
                        "{first_text}/{second_text} survivor {survivor}"
                    );
                }
            }
        }
    }

    #[test]
    fn paragraph_property_revisions_follow_second_mark_without_duplicate_ids() {
        let mut first = paragraph("00000001", "ab");
        let mut second = paragraph("00000002", "cd");
        first["propertyChanges"] = json!([{"type":"paragraphPropertyChange","info":{"id":1,"author":"Test","date":"2026-01-01T00:00:00Z"},"previousFormatting":{}}]);
        second["propertyChanges"] = json!([{"type":"paragraphPropertyChange","info":{"id":2,"author":"Test","date":"2026-01-01T00:00:00Z"},"previousFormatting":{}}]);
        let original = document(first, Some(second));
        for survivor in ["first", "second"] {
            let op = json!({"type":"joinBlocks","story":"main","blockId":"00000001","nextBlockId":"00000002","survivor":survivor});
            let changed = edit(&original, &op).unwrap();
            assert_eq!(
                changed.document["package"]["document"]["content"][0]["propertyChanges"][0]["info"]
                    ["id"],
                2
            );
            assert_eq!(apply_inverse(&changed.document, &changed.inverse), original);
        }
    }

    #[test]
    fn id_census_omits_derived_sections_and_includes_comment_paragraphs() {
        let mut original = document(paragraph("00000001", "ab"), None);
        original["package"]["document"]["sections"] =
            json!([{"properties":{},"content":original["package"]["document"]["content"]}]);
        let op = json!({"type":"splitBlock","at":{"story":"main","blockId":"00000001","offset":1},"newBlockId":"000000AB"});
        assert!(edit(&original, &op).is_ok());
        original["package"]["document"]["comments"] =
            json!([{"id":0,"author":"Test","content":[paragraph("000000ab","comment")]}]);
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "idCollision")
        );
    }
}
