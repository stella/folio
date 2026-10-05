//! Direct row and table operations. Row inverses restore `setTableRows`; an
//! empty row list removes the table and restores its whole containing block
//! list through `setContainerBlocks`, exactly as the TS structural primitive.

use crate::apply::{Applied, Path, Step, Touched, at_path, at_path_mut, locate, paragraph_paths};
use crate::refusal::Failure;
use serde_json::{Value, json};
use std::collections::{BTreeMap, BTreeSet};

struct TableLocation {
    table_path: Path,
    row_index: usize,
}

fn locate_table(document: &Value, op: &Value) -> Result<TableLocation, Failure> {
    let paragraph_path = locate(document, op)?;
    for length in (0..paragraph_path.len()).rev() {
        let path = &paragraph_path[..length];
        if at_path(document, path)
            .and_then(|value| value.get("type"))
            .and_then(Value::as_str)
            != Some("table")
        {
            continue;
        }
        let Some(Step::Field(field)) = paragraph_path.get(length) else {
            continue;
        };
        let Some(Step::Index(row_index)) = paragraph_path.get(length + 1) else {
            continue;
        };
        if field == "rows" {
            return Ok(TableLocation {
                table_path: path.to_vec(),
                row_index: *row_index,
            });
        }
    }
    Err(Failure::refused(
        op,
        "structureMismatch",
        "The paragraph is not in a table row.",
    ))
}

fn rows<'a>(table: &'a Value, op: &Value) -> Result<&'a Vec<Value>, Failure> {
    table
        .get("rows")
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "tableRowsShape"))
}

fn needs_table_edit(table: &Value) -> bool {
    ["preserved", "bookmarks", "carrierStack"]
        .iter()
        .any(|key| table.get(key).is_some())
        || table
            .get("rows")
            .and_then(Value::as_array)
            .is_some_and(|rows| {
                rows.iter().any(|row| {
                    row.get("contentControls").is_some() || row.get("carrierStack").is_some()
                })
            })
}

fn indexed_markup_refusal(op: &Value) -> Failure {
    Failure::refused(
        op,
        "untrackable",
        "Indexed table markup or shared row wrappers require a table operation.",
    )
}

/// The structural census in TS ids.ts sees all package stories, text boxes and
/// comments. Only the derived main-body section view is omitted.
fn paragraph_ids(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::Object(fields) => {
            if fields.get("type").and_then(Value::as_str) == Some("paragraph") {
                if let Some(id) = fields.get("paraId").and_then(Value::as_str) {
                    out.push(id.to_owned());
                }
            }
            for child in fields.values() {
                paragraph_ids(child, out);
            }
        }
        Value::Array(values) => {
            for child in values {
                paragraph_ids(child, out);
            }
        }
        _ => {}
    }
}

fn ids(value: &Value) -> Vec<String> {
    let mut out = Vec::new();
    paragraph_ids(value, &mut out);
    out
}

fn package_ids(document: &Value) -> Vec<String> {
    let mut out = Vec::new();
    let Some(package) = document.get("package").and_then(Value::as_object) else {
        return out;
    };
    for (key, value) in package {
        if key != "document" {
            paragraph_ids(value, &mut out);
            continue;
        }
        let Some(body) = value.as_object() else {
            continue;
        };
        for (field, child) in body {
            if field != "sections" {
                paragraph_ids(child, &mut out);
            }
        }
    }
    out
}

fn id_key(id: &str) -> String {
    id.to_ascii_uppercase()
}

fn valid_new_id(id: &str) -> bool {
    id.len() == 8
        && id.bytes().all(|byte| byte.is_ascii_hexdigit())
        && u32::from_str_radix(id, 16).is_ok_and(|number| number > 0 && number < 0x80000000)
}

fn story_paragraphs(value: &Value) -> Vec<&Value> {
    let mut paths = Vec::new();
    paragraph_paths(value, &mut Vec::new(), &mut paths);
    paths
        .iter()
        .filter_map(|path| at_path(value, path))
        .collect()
}

fn illegal_xml(value: &str) -> bool {
    value.chars().any(|character| {
        !matches!(character as u32,
        0x9 | 0xa | 0xd | 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
    })
}

fn inline_children(value: &Value) -> Option<&Vec<Value>> {
    let field = match value.get("type").and_then(Value::as_str) {
        Some("hyperlink") => "children",
        Some(
            "run" | "insertion" | "deletion" | "moveFrom" | "moveTo" | "inlineSdt"
            | "inlineWrapper",
        ) => "content",
        _ => return None,
    };
    value.get(field).and_then(Value::as_array)
}

fn illegal_text(value: &Value) -> bool {
    if value.get("type").and_then(Value::as_str) == Some("text") {
        return value
            .get("text")
            .and_then(Value::as_str)
            .is_some_and(illegal_xml);
    }
    inline_children(value).is_some_and(|children| children.iter().any(illegal_text))
}

fn empty_record(value: &Value) -> bool {
    let kind = value.get("type").and_then(Value::as_str);
    (kind == Some("text") && value.get("text").and_then(Value::as_str) == Some(""))
        || (matches!(
            kind,
            Some("run" | "insertion" | "deletion" | "moveFrom" | "moveTo")
        ) && value
            .get("content")
            .and_then(Value::as_array)
            .is_some_and(Vec::is_empty))
        || inline_children(value).is_some_and(|children| children.iter().any(empty_record))
}

/// Direct plain rows have no freshened revision/control identities. Carrying
/// one requires the full identity-space census and final-cell mark rules.
fn has_complex_identity(value: &Value) -> bool {
    match value {
        Value::Object(fields) => {
            matches!(
                fields.get("type").and_then(Value::as_str),
                Some("blockSdt" | "inlineSdt")
            ) || fields.get("info").is_some_and(|info| {
                info.get("id").is_some_and(Value::is_number)
                    && info.get("author").is_some_and(Value::is_string)
            }) || [
                "structuralChange",
                "contentControls",
                "carrierStack",
                "gridChange",
            ]
            .iter()
            .any(|field| fields.contains_key(*field))
                || fields.values().any(has_complex_identity)
        }
        Value::Array(values) => values.iter().any(has_complex_identity),
        _ => false,
    }
}

fn equal_for_staleness(left: &Value, right: &Value) -> bool {
    match (left, right) {
        (Value::Object(left), Value::Object(right)) => {
            let paragraph = left.get("type").and_then(Value::as_str) == Some("paragraph")
                && right.get("type").and_then(Value::as_str) == Some("paragraph");
            let skipped =
                |key: &str| paragraph && matches!(key, "listRendering" | "renderedPageBreakBefore");
            left.keys().filter(|key| !skipped(key)).count()
                == right.keys().filter(|key| !skipped(key)).count()
                && left
                    .iter()
                    .filter(|(key, _)| !skipped(key))
                    .all(|(key, value)| {
                        right
                            .get(key)
                            .is_some_and(|right| equal_for_staleness(value, right))
                    })
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| equal_for_staleness(left, right))
        }
        _ => left == right,
    }
}

fn validate_rows(value: &Value, op: &Value) -> Result<(), Failure> {
    let rows = value.as_array().expect("constructed row array");
    for row in rows {
        if row.get("type").and_then(Value::as_str) != Some("tableRow") {
            return Err(Failure::unsupported(op, "tableRowShape"));
        }
        let cells = row
            .get("cells")
            .and_then(Value::as_array)
            .ok_or_else(|| Failure::unsupported(op, "tableCellsShape"))?;
        if cells.is_empty() {
            return Err(Failure::refused(
                op,
                "structureMismatch",
                "Every row has cells, and every cell has a paragraph.",
            ));
        }
        for cell in cells {
            if cell.get("type").and_then(Value::as_str) != Some("tableCell") {
                return Err(Failure::unsupported(op, "tableCellShape"));
            }
            let blocks = cell
                .get("content")
                .filter(|content| content.is_array())
                .ok_or_else(|| Failure::unsupported(op, "tableCellContentShape"))?;
            if story_paragraphs(blocks).is_empty() {
                return Err(Failure::refused(
                    op,
                    "structureMismatch",
                    "Every row has cells, and every cell has a paragraph.",
                ));
            }
        }
    }
    let paragraphs = story_paragraphs(value);
    if paragraphs.iter().any(|paragraph| {
        paragraph
            .get("content")
            .and_then(Value::as_array)
            .is_some_and(|content| content.iter().any(illegal_text))
    }) {
        return Err(Failure::refused(
            op,
            "invalidText",
            "The row holds text that cannot be written to XML.",
        ));
    }
    if has_complex_identity(value) {
        return Err(Failure::unsupported(op, "rowRevisionAndControlIdentities"));
    }
    if paragraphs.iter().any(|paragraph| {
        paragraph
            .get("content")
            .and_then(Value::as_array)
            .is_some_and(|content| content.iter().any(empty_record))
    }) {
        return Err(Failure::refused(
            op,
            "emptyRecord",
            "A document story holds an empty run, text node, or revision wrapper; normalizeForOps removes them.",
        ));
    }
    Ok(())
}

/// A row inverse must address this table, never a nested table visited first.
fn table_anchor(rows: &Value) -> Option<&str> {
    fn in_blocks(blocks: &[Value]) -> Option<&str> {
        for block in blocks {
            match block.get("type").and_then(Value::as_str) {
                Some("paragraph") => return block.get("paraId").and_then(Value::as_str),
                Some("blockSdt" | "blockCustomXml") => {
                    if let Some(id) = block
                        .get("content")
                        .and_then(Value::as_array)
                        .and_then(|blocks| in_blocks(blocks))
                    {
                        return Some(id);
                    }
                }
                _ => {}
            }
        }
        None
    }
    for row in rows.as_array()? {
        for cell in row.get("cells")?.as_array()? {
            if let Some(id) = cell
                .get("content")
                .and_then(Value::as_array)
                .and_then(|blocks| in_blocks(blocks))
            {
                return Some(id);
            }
        }
    }
    None
}

fn touched(before: &Value, after: &Value) -> Touched {
    let before_ids = ids(before);
    let after_ids = ids(after);
    let before_set = before_ids
        .iter()
        .map(|id| id_key(id))
        .collect::<BTreeSet<_>>();
    let after_set = after_ids
        .iter()
        .map(|id| id_key(id))
        .collect::<BTreeSet<_>>();
    // TS reuses a whole row whenever it is structurally equal to any old row.
    // A changed row marks every retained paragraph, including unchanged text.
    let changed_ids = after
        .as_array()
        .expect("row list")
        .iter()
        .filter(|row| !before.as_array().expect("row list").contains(row))
        .flat_map(ids)
        .map(|id| id_key(&id))
        .collect::<BTreeSet<_>>();
    Touched {
        modified: before_ids
            .iter()
            .filter(|id| after_set.contains(&id_key(id)) && changed_ids.contains(&id_key(id)))
            .cloned()
            .collect(),
        inserted: after_ids
            .iter()
            .filter(|id| !before_set.contains(&id_key(id)))
            .cloned()
            .collect(),
        removed: before_ids
            .iter()
            .filter(|id| !after_set.contains(&id_key(id)))
            .cloned()
            .collect(),
    }
}

fn block_list_path(path: &[Step], op: &Value) -> Result<(Path, usize), Failure> {
    let Some((Step::Index(index), list)) = path.split_last() else {
        return Err(Failure::unsupported(op, "tableBlockListShape"));
    };
    Ok((list.to_vec(), *index))
}

fn structural_list<'a>(
    document: &'a Value,
    path: &[Step],
    op: &Value,
) -> Result<&'a Vec<Value>, Failure> {
    at_path(document, path)
        .and_then(Value::as_array)
        .ok_or_else(|| Failure::unsupported(op, "tableBlockListShape"))
}

/// Wrappers have container-final and identity rules outside this plain subset.
fn plain_structure(blocks: &[Value], op: &Value) -> Result<bool, Failure> {
    for block in blocks {
        match block.get("type").and_then(Value::as_str) {
            Some("paragraph") => {}
            Some("table") => {
                let rows = rows(block, op)?;
                if rows.is_empty() {
                    return Ok(false);
                }
                for row in rows {
                    if row.get("type").and_then(Value::as_str) != Some("tableRow") {
                        return Err(Failure::unsupported(op, "tableRowShape"));
                    }
                    let cells = row
                        .get("cells")
                        .and_then(Value::as_array)
                        .ok_or_else(|| Failure::unsupported(op, "tableCellsShape"))?;
                    if cells.is_empty() {
                        return Ok(false);
                    }
                    for cell in cells {
                        if cell.get("type").and_then(Value::as_str) != Some("tableCell") {
                            return Err(Failure::unsupported(op, "tableCellShape"));
                        }
                        let content = cell
                            .get("content")
                            .and_then(Value::as_array)
                            .ok_or_else(|| Failure::unsupported(op, "tableCellContentShape"))?;
                        if !plain_structure(content, op)? || !ends_in_paragraph(content) {
                            return Ok(false);
                        }
                    }
                }
            }
            _ => {
                return Err(Failure::unsupported(
                    op,
                    "tableStructuralBlockWrappersAndMarkers",
                ));
            }
        }
    }
    Ok(true)
}

fn ends_in_paragraph(blocks: &[Value]) -> bool {
    blocks
        .last()
        .and_then(|block| block.get("type"))
        .and_then(Value::as_str)
        == Some("paragraph")
}

fn plain_container(document: &Value, path: &[Step], op: &Value) -> Result<(), Failure> {
    for length in 0..path.len() {
        if matches!(
            at_path(document, &path[..length])
                .and_then(|value| value.get("type"))
                .and_then(Value::as_str),
            Some("blockSdt" | "blockCustomXml")
        ) {
            return Err(Failure::unsupported(
                op,
                "tableStructuralBlockWrappersAndMarkers",
            ));
        }
    }
    Ok(())
}

fn paragraph_map<'a>(paragraphs: &[&'a Value]) -> BTreeMap<String, &'a Value> {
    paragraphs
        .iter()
        .filter_map(|paragraph| {
            paragraph
                .get("paraId")
                .and_then(Value::as_str)
                .map(|id| (id_key(id), *paragraph))
        })
        .collect()
}

fn commit_blocks(
    document: &Value,
    op: &Value,
    anchor_path: &[Step],
    blocks: Vec<Value>,
) -> Result<Applied, Failure> {
    let (list_path, _) = block_list_path(anchor_path, op)?;
    let before = structural_list(document, &list_path, op)?;
    if *before == blocks {
        return Ok(Applied {
            document: document.clone(),
            inverse: vec![],
            touched: Touched::default(),
            revisions: vec![],
        });
    }
    let anchor = at_path(document, anchor_path).expect("located structural anchor");
    let anchor_id = anchor
        .get("paraId")
        .and_then(Value::as_str)
        .ok_or_else(|| Failure::unsupported(op, "paragraphIdShape"))?;
    if !blocks.iter().any(|block| {
        block.get("type").and_then(Value::as_str) == Some("paragraph")
            && block
                .get("paraId")
                .and_then(Value::as_str)
                .is_some_and(|id| id.eq_ignore_ascii_case(anchor_id))
    }) {
        return Err(Failure::refused(
            op,
            "untrackable",
            "The structural inverse requires a surviving paragraph in this block list.",
        ));
    }
    plain_container(document, &list_path, op)?;
    if blocks.is_empty()
        || !plain_structure(before, op)?
        || !plain_structure(&blocks, op)?
        || !ends_in_paragraph(before)
        || !ends_in_paragraph(&blocks)
    {
        return Err(Failure::refused(
            op,
            "structureMismatch",
            "A table has rows and cells; each cell and the story retain a final paragraph.",
        ));
    }
    let before_value = Value::Array(before.clone());
    let after = Value::Array(blocks);
    let before_paragraphs = story_paragraphs(&before_value);
    let after_paragraphs = story_paragraphs(&after);
    let before_by_id = paragraph_map(&before_paragraphs);
    let after_by_id = paragraph_map(&after_paragraphs);
    for paragraph in before_paragraphs.iter().chain(&after_paragraphs) {
        let id = paragraph.get("paraId").and_then(Value::as_str);
        let key = id_key(id.unwrap_or(""));
        let opposite = if before_by_id.contains_key(&key) {
            &after_by_id
        } else {
            &before_by_id
        };
        if paragraph.get("sectionProperties")
            != opposite
                .get(&key)
                .and_then(|paragraph| paragraph.get("sectionProperties"))
        {
            return Err(Failure::refused(
                op,
                "sectionBoundary",
                "Table edits preserve section boundaries.",
            ));
        }
        if id.is_none_or(|id| {
            (!before_by_id.contains_key(&key) || !after_by_id.contains_key(&key))
                && !valid_new_id(id)
        }) {
            return Err(Failure::refused(
                op,
                "invalidBlockId",
                "An added or removed paragraph needs a usable id.",
            ));
        }
        let content = paragraph
            .get("content")
            .and_then(Value::as_array)
            .ok_or_else(|| Failure::unsupported(op, "paragraphContentShape"))?;
        if content.iter().any(illegal_text) {
            return Err(Failure::refused(
                op,
                "invalidText",
                "The table holds text that cannot be written to XML.",
            ));
        }
        if content.iter().any(empty_record) {
            return Err(Failure::refused(
                op,
                "emptyRecord",
                "A document story holds an empty run, text node, or revision wrapper; normalizeForOps removes them.",
            ));
        }
        if paragraph.get("pPrMark").is_some() {
            return Err(Failure::unsupported(
                op,
                "tableStructuralFinalParagraphMarks",
            ));
        }
        crate::inline::validate_plain(content, op, false)?;
    }
    if has_complex_identity(&before_value) || has_complex_identity(&after) {
        return Err(Failure::unsupported(
            op,
            "tableStructuralRevisionAndControlIdentities",
        ));
    }
    let mut remaining = BTreeMap::<String, usize>::new();
    for id in package_ids(document) {
        *remaining.entry(id_key(&id)).or_default() += 1;
    }
    for id in ids(&before_value) {
        let count = remaining
            .get_mut(&id_key(&id))
            .expect("container paragraph counted in its package");
        *count -= 1;
    }
    let mut incoming = BTreeSet::new();
    for id in ids(&after) {
        let key = id_key(&id);
        if remaining.get(&key).is_some_and(|count| *count > 0) || !incoming.insert(key) {
            return Err(Failure::refused(
                op,
                "idCollision",
                "A paragraph id is already used in the package.",
            ));
        }
    }
    let change = Touched {
        modified: after_paragraphs
            .iter()
            .filter(|paragraph| {
                let key = id_key(
                    paragraph["paraId"]
                        .as_str()
                        .expect("validated paragraph id"),
                );
                before_by_id
                    .get(&key)
                    .is_some_and(|before| *before != **paragraph)
            })
            .map(|paragraph| {
                paragraph["paraId"]
                    .as_str()
                    .expect("validated paragraph id")
                    .to_owned()
            })
            .collect(),
        inserted: after_paragraphs
            .iter()
            .filter(|paragraph| {
                !before_by_id.contains_key(&id_key(
                    paragraph["paraId"]
                        .as_str()
                        .expect("validated paragraph id"),
                ))
            })
            .map(|paragraph| {
                paragraph["paraId"]
                    .as_str()
                    .expect("validated paragraph id")
                    .to_owned()
            })
            .collect(),
        removed: before_paragraphs
            .iter()
            .filter(|paragraph| {
                !after_by_id.contains_key(&id_key(
                    paragraph["paraId"]
                        .as_str()
                        .expect("validated paragraph id"),
                ))
            })
            .map(|paragraph| {
                paragraph["paraId"]
                    .as_str()
                    .expect("validated paragraph id")
                    .to_owned()
            })
            .collect(),
    };
    let inverse = json!({"type":"setContainerBlocks","story":op["story"],"blockId":anchor_id,"expected":after,"blocks":before_value});
    let mut result = document.clone();
    *at_path_mut(&mut result, &list_path).expect("located mutable block list") = after;
    Ok(Applied {
        document: result,
        inverse: vec![inverse],
        touched: change,
        revisions: vec![],
    })
}

fn delete_table(document: &Value, op: &Value) -> Result<Applied, Failure> {
    let location = locate_table(document, op)?;
    let table = at_path(document, &location.table_path).expect("located table");
    if op
        .get("expected")
        .is_some_and(|expected| !equal_for_staleness(table, expected))
    {
        return Err(Failure::refused(
            op,
            "stale",
            "The table to remove has changed.",
        ));
    }
    let (list_path, index) = block_list_path(&location.table_path, op)?;
    let before = structural_list(document, &list_path, op)?;
    let anchor_index = before
        .iter()
        .position(|block| block.get("type").and_then(Value::as_str) == Some("paragraph"))
        .ok_or_else(|| {
            Failure::refused(
                op,
                "untrackable",
                "Removing a table requires a surviving paragraph in its block list.",
            )
        })?;
    let mut anchor_path = list_path;
    anchor_path.push(Step::Index(anchor_index));
    let mut blocks = before.clone();
    blocks.remove(index);
    commit_blocks(document, op, &anchor_path, blocks)
}

fn structural_edit(document: &Value, op: &Value) -> Result<Applied, Failure> {
    match op.get("type").and_then(Value::as_str) {
        Some("deleteTable") => delete_table(document, op),
        Some("setContainerBlocks") => {
            let anchor_path = locate(document, op)?;
            let (list_path, _) = block_list_path(&anchor_path, op)?;
            let before = structural_list(document, &list_path, op)?;
            let expected = op
                .get("expected")
                .filter(|value| value.is_array())
                .ok_or_else(|| {
                    Failure::refused(
                        op,
                        "invalidOperation",
                        "A container restoration must name its expected blocks.",
                    )
                })?;
            if !equal_for_staleness(&Value::Array(before.clone()), expected) {
                return Err(Failure::refused(
                    op,
                    "stale",
                    "The container blocks have changed.",
                ));
            }
            let blocks = op.get("blocks").and_then(Value::as_array).ok_or_else(|| {
                Failure::refused(
                    op,
                    "invalidOperation",
                    "A container restoration must name its blocks.",
                )
            })?;
            commit_blocks(document, op, &anchor_path, blocks.clone())
        }
        Some("insertTable") => {
            let at = op.get("at").ok_or_else(|| {
                Failure::refused(
                    op,
                    "invalidOperation",
                    "A table insertion must name its anchor.",
                )
            })?;
            let direction = at.get("type").and_then(Value::as_str);
            if !matches!(direction, Some("before" | "after")) {
                return Err(Failure::refused(
                    op,
                    "structureMismatch",
                    "Insert before or after a paragraph.",
                ));
            }
            let locator = json!({"type":op["type"],"story":op["story"],"blockId":at["blockId"]});
            let anchor_path = locate(document, &locator).map_err(|failure| failure.with_op(op))?;
            let (list_path, index) = block_list_path(&anchor_path, op)?;
            plain_container(document, &list_path, op)?;
            let before = structural_list(document, &list_path, op)?;
            let terminal = direction == Some("after") && index + 1 == before.len();
            if op.get("terminal").is_some() && !terminal {
                return Err(Failure::refused(
                    op,
                    "structureMismatch",
                    "A terminal insertion must follow its container's final paragraph.",
                ));
            }
            if terminal && op.get("terminal").is_none() {
                return Err(Failure::refused(
                    op,
                    "structureMismatch",
                    "A terminal table insertion supplies the preceding paragraph id.",
                ));
            }
            let table = op
                .get("table")
                .filter(|value| value.get("type").and_then(Value::as_str) == Some("table"))
                .ok_or_else(|| Failure::unsupported(op, "tableShape"))?;
            let mut blocks = before.clone();
            if let Some(terminal) = op.get("terminal") {
                let anchor = &before[index];
                let mut preceding = json!({"type":"paragraph","paraId":terminal["beforeBlockId"],"content":anchor["content"]});
                if let Some(formatting) = anchor.get("formatting") {
                    preceding["formatting"] = formatting.clone();
                }
                let mut carrier = anchor.clone();
                carrier["content"] = json!([]);
                blocks.splice(index..=index, [preceding, table.clone(), carrier]);
            } else {
                blocks.insert(
                    index + usize::from(direction == Some("after")),
                    table.clone(),
                );
            }
            commit_blocks(document, op, &anchor_path, blocks)
        }
        _ => Err(Failure::unsupported(op, "tableStructuralOperationKind")),
    }
}

fn commit(
    document: &Value,
    op: &Value,
    location: &TableLocation,
    replacement: Vec<Value>,
) -> Result<Applied, Failure> {
    let table = at_path(document, &location.table_path).expect("located table");
    let before = &table["rows"];
    let after = Value::Array(replacement);
    if after.as_array().expect("row list").is_empty() {
        let deletion = json!({"type":"deleteTable","story":op["story"],"blockId":op["blockId"],"expected":table});
        return delete_table(document, &deletion).map_err(|failure| failure.with_op(op));
    }
    if *before == after {
        return Ok(Applied {
            document: document.clone(),
            inverse: vec![],
            touched: Touched::default(),
            revisions: vec![],
        });
    }
    validate_rows(&after, op)?;
    let original_ids = ids(before)
        .into_iter()
        .map(|id| id_key(&id))
        .collect::<BTreeSet<_>>();
    for paragraph in story_paragraphs(&after) {
        let id = paragraph.get("paraId").and_then(Value::as_str);
        if id.is_none_or(|id| !original_ids.contains(&id_key(id)) && !valid_new_id(id)) {
            return Err(Failure::refused(
                op,
                "invalidBlockId",
                "An added paragraph has no usable id.",
            ));
        }
    }
    let mut remaining = BTreeMap::<String, usize>::new();
    for id in package_ids(document) {
        *remaining.entry(id_key(&id)).or_default() += 1;
    }
    for id in ids(before) {
        let count = remaining
            .get_mut(&id_key(&id))
            .expect("table paragraph counted in its package");
        *count -= 1;
    }
    let mut incoming = BTreeSet::new();
    for id in ids(&after) {
        let key = id_key(&id);
        if remaining.get(&key).is_some_and(|count| *count > 0) || !incoming.insert(key) {
            return Err(Failure::refused(
                op,
                "idCollision",
                "A row paragraph id is already used in the package.",
            ));
        }
    }
    let anchor = table_anchor(&after).ok_or_else(|| {
        Failure::refused(
            op,
            "untrackable",
            "The table has no paragraph that can address its rows.",
        )
    })?;
    let inverse = json!({"type":"setTableRows","story":op["story"],"blockId":anchor,"expected":after,"rows":before});
    let change = touched(before, &after);
    let mut result = document.clone();
    at_path_mut(&mut result, &location.table_path).expect("located mutable table")["rows"] = after;
    Ok(Applied {
        document: result,
        inverse: vec![inverse],
        touched: change,
        revisions: vec![],
    })
}

pub fn edit(document: &Value, op: &Value) -> Result<Applied, Failure> {
    if op.get("revision").is_some() {
        return Err(Failure::unsupported(op, "trackedTableRows"));
    }
    if matches!(
        op.get("type").and_then(Value::as_str),
        Some("deleteTable" | "insertTable" | "setContainerBlocks")
    ) {
        return structural_edit(document, op);
    }
    let location = locate_table(document, op)?;
    let table = at_path(document, &location.table_path).expect("located table");
    let before = rows(table, op)?;
    let mut replacement = before.clone();
    match op.get("type").and_then(Value::as_str) {
        Some("insertRow") => {
            let at = op
                .get("at")
                .and_then(Value::as_u64)
                .and_then(|at| usize::try_from(at).ok())
                .filter(|at| *at <= before.len())
                .ok_or_else(|| {
                    Failure::refused(op, "invalidOffset", "The row index is outside the table.")
                })?;
            if needs_table_edit(table) {
                return Err(indexed_markup_refusal(op));
            }
            let row = op.get("row").filter(|row| row.is_object()).ok_or_else(|| {
                Failure::refused(op, "invalidOperation", "An insertion must name its row.")
            })?;
            let existing = ids(&table["rows"])
                .into_iter()
                .map(|id| id_key(&id))
                .collect::<BTreeSet<_>>();
            if ids(row).iter().any(|id| existing.contains(&id_key(id))) {
                return Err(Failure::refused(
                    op,
                    "idCollision",
                    "An inserted paragraph id is already used in the table.",
                ));
            }
            replacement.insert(at, row.clone());
        }
        Some("deleteRow") => {
            let row = before.get(location.row_index).expect("located row");
            if op
                .get("expected")
                .is_some_and(|expected| !equal_for_staleness(row, expected))
            {
                return Err(Failure::refused(
                    op,
                    "stale",
                    "The row to remove has changed.",
                ));
            }
            if before.len() == 1 {
                return commit(document, op, &location, vec![]);
            }
            if needs_table_edit(table) {
                return Err(indexed_markup_refusal(op));
            }
            replacement.remove(location.row_index);
        }
        Some("setTableRows") => {
            let expected = op
                .get("expected")
                .filter(|expected| expected.is_array())
                .ok_or_else(|| {
                    Failure::refused(
                        op,
                        "invalidOperation",
                        "A row restoration must name its expected rows.",
                    )
                })?;
            if !equal_for_staleness(&table["rows"], expected) {
                return Err(Failure::refused(
                    op,
                    "stale",
                    "The table rows have changed.",
                ));
            }
            let incoming = op.get("rows").and_then(Value::as_array).ok_or_else(|| {
                Failure::refused(
                    op,
                    "invalidOperation",
                    "A row restoration must name its rows.",
                )
            })?;
            if !incoming.is_empty() && incoming.len() != before.len() && needs_table_edit(table) {
                return Err(indexed_markup_refusal(op));
            }
            replacement = incoming.clone();
        }
        _ => return Err(Failure::unsupported(op, "tableRowOperationKind")),
    }
    if !replacement.is_empty() && replacement != *before && has_complex_identity(&table["rows"]) {
        return Err(Failure::unsupported(op, "rowRevisionAndControlIdentities"));
    }
    commit(document, op, &location, replacement)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph(id: &str) -> Value {
        json!({"type":"paragraph","paraId":id,"content":[{"type":"run","content":[{"type":"text","text":id}]}]})
    }

    fn row(id: &str) -> Value {
        json!({"type":"tableRow","preservedAttributes":[{"name":"rsidR","value":"abc"}],
            "cells":[{"type":"tableCell","id":"authored-cell","content":[paragraph(id)]}]})
    }

    fn document() -> Value {
        json!({"package":{"document":{"content":[{"type":"table","unknown":{"rawXml":"<w:custom/>"},
            "rows":[row("00000001"),row("00000002")]},paragraph("00000003")]}}})
    }

    fn locator(kind: &str, id: &str) -> Value {
        json!({"type":kind,"story":"main","blockId":id})
    }

    #[test]
    fn insert_and_delete_at_every_row_index_restore_exact_authored_records() {
        let original = document();
        for at in 0..=2 {
            let mut op = locator("insertRow", "00000001");
            op["at"] = json!(at);
            op["row"] = row("00000004");
            let changed = edit(&original, &op).unwrap();
            assert_eq!(changed.touched.inserted, ["00000004"]);
            assert!(changed.touched.modified.is_empty());
            let restored = edit(&changed.document, &changed.inverse[0]).unwrap();
            assert_eq!(restored.document, original);
        }
        for id in ["00000001", "00000002"] {
            let changed = edit(&original, &locator("deleteRow", id)).unwrap();
            assert_eq!(changed.touched.removed, [id]);
            let restored = edit(&changed.document, &changed.inverse[0]).unwrap();
            assert_eq!(restored.document, original);
        }
    }

    #[test]
    fn changed_row_metadata_marks_retained_paragraphs_but_reorders_reuse_rows() {
        let original = document();
        let before = original["package"]["document"]["content"][0]["rows"].clone();
        let mut incoming = before.clone();
        incoming[0]["formatting"] = json!({"cantSplit":true});
        let mut op = locator("setTableRows", "00000001");
        op["expected"] = before.clone();
        op["rows"] = incoming;
        let changed = edit(&original, &op).unwrap();
        assert_eq!(changed.touched.modified, ["00000001"]);
        assert_eq!(
            edit(&changed.document, &changed.inverse[0])
                .unwrap()
                .document,
            original
        );
        op["rows"] = json!([before[1], before[0]]);
        let reordered = edit(&original, &op).unwrap();
        assert!(reordered.touched.modified.is_empty());
        assert!(reordered.touched.inserted.is_empty());
        assert!(reordered.touched.removed.is_empty());
    }

    #[test]
    fn nested_paragraph_selects_innermost_table_and_inverse_avoids_nested_anchor() {
        let mut original = document();
        let nested = json!({"type":"table","rows":[row("00000005"),row("00000006")]});
        original["package"]["document"]["content"][0]["rows"][0]["cells"][0]["content"] =
            json!([nested, paragraph("00000001")]);
        let changed = edit(&original, &locator("deleteRow", "00000005")).unwrap();
        assert_eq!(changed.touched.removed, ["00000005"]);
        assert_eq!(changed.inverse[0]["blockId"], "00000006");
        assert_eq!(
            edit(&changed.document, &changed.inverse[0])
                .unwrap()
                .document,
            original
        );
        let outer = edit(&original, &locator("deleteRow", "00000002")).unwrap();
        assert_eq!(outer.inverse[0]["blockId"], "00000001");
        assert_eq!(
            edit(&outer.document, &outer.inverse[0]).unwrap().document,
            original
        );
    }

    #[test]
    fn census_ignores_sections_but_checks_ids_in_other_package_parts_case_insensitively() {
        let mut original = document();
        original["package"]["document"]["sections"] =
            json!([{"properties":{},"content":original["package"]["document"]["content"]}]);
        let mut op = locator("insertRow", "00000001");
        op["at"] = json!(1);
        op["row"] = row("000000AB");
        assert!(edit(&original, &op).is_ok());
        original["package"]["document"]["comments"] =
            json!([{"id":0,"author":"Test","content":[paragraph("000000ab")]}]);
        assert!(
            matches!(edit(&original,&op),Err(Failure::Refused {reason,..}) if reason == "idCollision")
        );
    }

    #[test]
    fn every_empty_row_path_restores_the_exact_container_and_table_metadata() {
        for table_index in 0..3 {
            for kind in ["deleteRow", "setTableRows", "deleteTable"] {
                let table = json!({"type":"table","rows":[row("00000001")],
                    "preserved":[{"rawXml":"<w:custom/>","position":0}],
                    "unknown":{"rawXml":"<w:opaque/>"},"formatting":{"styleId":"Authored"}});
                let mut blocks = vec![
                    paragraph("00000002"),
                    paragraph("00000003"),
                    paragraph("00000004"),
                ];
                blocks[0]["authored"] = json!({"opaque":[1,2,3]});
                blocks.insert(table_index, table.clone());
                let original = json!({"package":{"document":{"content":blocks,"customBodyField":{"keep":true}}}});
                let mut op = locator(kind, "00000001");
                match kind {
                    "deleteRow" => op["expected"] = table["rows"][0].clone(),
                    "setTableRows" => {
                        op["expected"] = table["rows"].clone();
                        op["rows"] = json!([]);
                    }
                    "deleteTable" => op["expected"] = table.clone(),
                    _ => unreachable!(),
                }
                let changed = edit(&original, &op).unwrap();
                let expected = json!([
                    paragraph("00000002"),
                    paragraph("00000003"),
                    paragraph("00000004")
                ]);
                let mut expected = expected;
                expected[0]["authored"] = json!({"opaque":[1,2,3]});
                assert_eq!(changed.document["package"]["document"]["content"], expected);
                assert_eq!(
                    changed.inverse,
                    [
                        json!({"type":"setContainerBlocks","story":"main","blockId":"00000002",
                    "expected":expected,"blocks":original["package"]["document"]["content"]})
                    ]
                );
                assert_eq!(changed.touched.removed, ["00000001"]);
                assert!(changed.touched.modified.is_empty());
                assert!(changed.touched.inserted.is_empty());
                let restored = edit(&changed.document, &changed.inverse[0]).unwrap();
                assert_eq!(restored.document, original);
                assert_eq!(restored.touched.inserted, ["00000001"]);
                assert_eq!(
                    edit(&restored.document, &restored.inverse[0])
                        .unwrap()
                        .document,
                    changed.document
                );
            }
        }
    }

    #[test]
    fn deletion_anchor_survives_in_the_same_nested_cell_and_not_another_list() {
        let nested = json!({"type":"table","rows":[row("00000005")]});
        let mut original = document();
        original["package"]["document"]["content"][0]["rows"][0]["cells"][0]["content"] =
            json!([nested, paragraph("00000001")]);
        let changed = edit(&original, &locator("deleteRow", "00000005")).unwrap();
        assert_eq!(changed.inverse[0]["blockId"], "00000001");
        assert_eq!(
            changed.inverse[0]["expected"],
            json!([paragraph("00000001")])
        );
        assert_eq!(
            edit(&changed.document, &changed.inverse[0])
                .unwrap()
                .document,
            original
        );
        original["package"]["document"]["content"][0]["rows"][0]["cells"][0]["content"] =
            json!([{"type":"table","rows":[row("00000005")]}]);
        assert!(
            matches!(edit(&original, &locator("deleteRow","00000005")), Err(Failure::Refused {reason,..}) if reason == "untrackable")
        );
    }

    #[test]
    fn container_inverse_staleness_covers_every_sibling_but_ignores_rendered_fields() {
        let mut original = document();
        original["package"]["document"]["content"][0]["rows"] = json!([row("00000001")]);
        let changed = edit(&original, &locator("deleteRow", "00000001")).unwrap();
        let mut stale = changed.document.clone();
        stale["package"]["document"]["content"][0]["formatting"] = json!({"keepNext":true});
        assert!(
            matches!(edit(&stale, &changed.inverse[0]), Err(Failure::Refused {reason,..}) if reason == "stale")
        );
        let mut rendered = changed.document.clone();
        rendered["package"]["document"]["content"][0]["listRendering"] = json!({"label":"1."});
        assert_eq!(
            edit(&rendered, &changed.inverse[0]).unwrap().document,
            original
        );
    }

    #[test]
    fn table_insertion_at_every_anchor_has_a_whole_container_inverse() {
        let original = json!({"package":{"document":{"content":[paragraph("00000001"),paragraph("00000002"),paragraph("00000003")]}}});
        let table = json!({"type":"table","unknown":{"authored":true},"rows":[row("00000004")]});
        for id in ["00000001", "00000002", "00000003"] {
            for direction in ["before", "after"] {
                let mut op = json!({"type":"insertTable","story":"main","at":{"type":direction,"blockId":id},"table":table});
                if id == "00000003" && direction == "after" {
                    op["terminal"] = json!({"beforeBlockId":"00000005"});
                }
                let changed = edit(&original, &op).unwrap();
                assert_eq!(changed.inverse[0]["type"], "setContainerBlocks");
                assert_eq!(changed.inverse[0]["blockId"], id);
                assert_eq!(
                    changed.inverse[0]["blocks"],
                    original["package"]["document"]["content"]
                );
                assert_eq!(
                    changed.inverse[0]["expected"],
                    changed.document["package"]["document"]["content"]
                );
                assert_eq!(
                    edit(&changed.document, &changed.inverse[0])
                        .unwrap()
                        .document,
                    original
                );
                if id == "00000003" && direction == "after" {
                    assert_eq!(changed.touched.modified, ["00000003"]);
                    assert_eq!(changed.touched.inserted, ["00000005", "00000004"]);
                } else {
                    assert!(changed.touched.modified.is_empty());
                    assert_eq!(changed.touched.inserted, ["00000004"]);
                }
            }
        }
    }

    #[test]
    fn terminal_insertion_preserves_carrier_metadata_and_copies_only_preceding_format_and_content()
    {
        let mut carrier = paragraph("00000001");
        carrier["formatting"] = json!({"keepNext":true});
        carrier["sectionProperties"] = json!({"type":"nextPage"});
        carrier["unknown"] = json!({"keep":"on-carrier"});
        let original = json!({"package":{"document":{"content":[carrier]}}});
        let table = json!({"type":"table","rows":[row("00000002")]});
        let op = json!({"type":"insertTable","story":"main","at":{"type":"after","blockId":"00000001"},
            "table":table,"terminal":{"beforeBlockId":"00000003"}});
        let changed = edit(&original, &op).unwrap();
        let blocks = &changed.document["package"]["document"]["content"];
        assert_eq!(
            blocks[0],
            json!({"type":"paragraph","paraId":"00000003","formatting":{"keepNext":true},"content":carrier["content"]})
        );
        let mut empty_carrier = carrier;
        empty_carrier["content"] = json!([]);
        assert_eq!(blocks[1], table);
        assert_eq!(blocks[2], empty_carrier);
        assert_eq!(
            edit(&changed.document, &changed.inverse[0])
                .unwrap()
                .document,
            original
        );
    }

    #[test]
    fn structural_paths_refuse_unrestorable_boundaries_and_package_collisions() {
        let mut original = document();
        original["package"]["document"]["content"][0]["rows"] = json!([row("00000001")]);
        original["package"]["document"]["content"][0]["rows"][0]["cells"][0]["content"][0]["sectionProperties"] =
            json!({});
        assert!(
            matches!(edit(&original, &locator("deleteRow","00000001")), Err(Failure::Refused {reason,..}) if reason == "sectionBoundary")
        );
        let original = document();
        let op = json!({"type":"insertTable","story":"main","at":{"type":"before","blockId":"00000003"},
            "table":{"type":"table","rows":[row("00000001")]}});
        assert!(
            matches!(edit(&original, &op), Err(Failure::Refused {reason,..}) if reason == "idCollision")
        );
        let mut missing_terminal = op;
        missing_terminal["at"]["type"] = json!("after");
        assert!(
            matches!(edit(&original, &missing_terminal), Err(Failure::Refused {reason,..}) if reason == "structureMismatch")
        );
    }

    #[test]
    fn unsupported_structural_identity_dimensions_do_not_escape_as_partial_edits() {
        let mut original = document();
        original["package"]["document"]["content"][0]["rows"] = json!([row("00000001")]);
        original["package"]["document"]["content"][0]["rows"][0]["structuralChange"] =
            json!({"kind":"ins","info":{"id":1,"author":"Test"}});
        assert!(
            matches!(edit(&original, &locator("deleteRow","00000001")), Err(Failure::Unsupported {dimension,..}) if dimension == "tableStructuralRevisionAndControlIdentities")
        );
    }
}
