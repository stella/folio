//! Resident JSON model and experiment-only main-content projection patches.
//! Package changes need a wider patch contract and are explicitly unsupported.
//! Parsing and full save are separate from operation dispatch and patch transfer.

use crate::apply::{Applied, Touched, apply_operations};
use crate::refusal::Failure;
use serde::Serialize;
use serde_json::{Map, Value};
use std::collections::BTreeMap;

pub struct ResidentDocument {
    value: Value,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ProjectionPatch {
    /// A number references the prior top-level content; an object is a new block.
    content: Vec<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    section_metadata: Option<Vec<Value>>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ResidentApplied {
    inverse: Vec<Value>,
    touched: Touched,
    revisions: Vec<i64>,
    projection_patch: ProjectionPatch,
}

fn main_content(document: &Value) -> Option<&Vec<Value>> {
    document
        .pointer("/package/document/content")
        .and_then(Value::as_array)
}

fn same_fields_except(
    before: &Map<String, Value>,
    after: &Map<String, Value>,
    omitted: &[&str],
) -> bool {
    let retained = |key: &String| !omitted.contains(&key.as_str());
    before.keys().filter(|key| retained(key)).count()
        == after.keys().filter(|key| retained(key)).count()
        && before
            .iter()
            .filter(|(key, _)| retained(key))
            .all(|(key, value)| after.get(key) == Some(value))
}

/// Compare resident package facts without cloning or serializing main content.
fn package_facts_unchanged(before: &Value, after: &Value) -> bool {
    let (Some(old_root), Some(new_root)) = (before.as_object(), after.as_object()) else {
        return false;
    };
    if !same_fields_except(old_root, new_root, &["package"]) {
        return false;
    }
    let (Some(old_package), Some(new_package)) =
        (before["package"].as_object(), after["package"].as_object())
    else {
        return false;
    };
    if !same_fields_except(old_package, new_package, &["document"]) {
        return false;
    }
    let (Some(old_body), Some(new_body)) = (
        before["package"]["document"].as_object(),
        after["package"]["document"].as_object(),
    ) else {
        return false;
    };
    same_fields_except(old_body, new_body, &["content", "sections"])
}

fn section_metadata(document: &Value, op: &Value) -> Result<Option<Vec<Value>>, Failure> {
    let Some(value) = document.pointer("/package/document/sections") else {
        return Ok(None);
    };
    let sections = value
        .as_array()
        .ok_or_else(|| Failure::unsupported(op, "residentSectionProjection"))?;
    let mut metadata = Vec::new();
    for section in sections {
        let mut fields = section
            .as_object()
            .cloned()
            .ok_or_else(|| Failure::unsupported(op, "residentSectionProjection"))?;
        fields.remove("content");
        metadata.push(Value::Object(fields));
    }
    Ok(Some(metadata))
}

fn projection_patch(before: &Value, after: &Value, op: &Value) -> Result<ProjectionPatch, Failure> {
    let old =
        main_content(before).ok_or_else(|| Failure::unsupported(op, "residentDocumentShape"))?;
    let new =
        main_content(after).ok_or_else(|| Failure::unsupported(op, "residentDocumentShape"))?;
    if !package_facts_unchanged(before, after) {
        return Err(Failure::unsupported(op, "residentPackageProjection"));
    }
    // An omitted optional metadata field cannot express removal of owned sections.
    if before.pointer("/package/document/sections").is_some()
        && after.pointer("/package/document/sections").is_none()
    {
        return Err(Failure::unsupported(op, "residentSectionProjection"));
    }
    let mut paragraph_indices = BTreeMap::new();
    for (index, block) in old.iter().enumerate() {
        if let Some(id) = block.get("paraId").and_then(Value::as_str) {
            paragraph_indices.insert(id, index);
        }
    }
    let mut content = Vec::with_capacity(new.len());
    for (index, block) in new.iter().enumerate() {
        let same_index = old
            .get(index)
            .filter(|candidate| *candidate == block)
            .map(|_| index);
        let same_paragraph = block
            .get("paraId")
            .and_then(Value::as_str)
            .and_then(|id| paragraph_indices.get(id))
            .copied()
            .filter(|index| old.get(*index) == Some(block));
        let prior = same_index.or(same_paragraph).or_else(|| {
            // Tables and other block records do not own a paragraph id.
            if block.get("paraId").is_some() {
                return None;
            }
            old.iter().position(|candidate| candidate == block)
        });
        match prior {
            Some(index) => content.push(Value::from(index)),
            None if block.is_object() => content.push(block.clone()),
            None => return Err(Failure::unsupported(op, "residentBlockProjection")),
        }
    }
    Ok(ProjectionPatch {
        content,
        section_metadata: section_metadata(after, op)?,
    })
}

impl ResidentDocument {
    pub fn load(document_json: &str) -> Result<Self, String> {
        let value: Value =
            serde_json::from_str(document_json).map_err(|error| error.to_string())?;
        if main_content(&value).is_none() {
            return Err("A resident document needs package.document.content as an array.".into());
        }
        Ok(Self { value })
    }

    pub fn save(&self) -> Result<String, String> {
        serde_json::to_string(&self.value).map_err(|error| error.to_string())
    }

    /// Apply atomically. A semantic failure is returned as serialized Failure;
    /// JSON parse/serialization failures use the outer transport error channel.
    pub fn apply_ops_json(&mut self, operations_json: &str) -> Result<String, String> {
        let operations: Vec<Value> =
            serde_json::from_str(operations_json).map_err(|error| error.to_string())?;
        if let Some(op) = operations.iter().find(|op| {
            matches!(
                op["type"].as_str(),
                Some(
                    "createComment"
                        | "updateCommentContent"
                        | "setCommentResolution"
                        | "deleteComment"
                        | "restoreCommentState"
                )
            )
        }) {
            return serde_json::to_string(&Failure::unsupported(op, "residentPackageProjection"))
                .map_err(|error| error.to_string());
        }
        let applied = match apply_operations(&self.value, &operations) {
            Ok(applied) => applied,
            Err(failure) => {
                return serde_json::to_string(&failure).map_err(|error| error.to_string());
            }
        };
        let empty_operation = Value::Null;
        let op = operations.last().unwrap_or(&empty_operation);
        let patch = match projection_patch(&self.value, &applied.document, op) {
            Ok(patch) => patch,
            Err(failure) => {
                return serde_json::to_string(&failure).map_err(|error| error.to_string());
            }
        };
        let Applied {
            document,
            inverse,
            touched,
            revisions,
        } = applied;
        let output = serde_json::to_string(&ResidentApplied {
            inverse,
            touched,
            revisions,
            projection_patch: patch,
        })
        .map_err(|error| error.to_string())?;
        // No parse, application, patch refusal, or serialization error can commit.
        self.value = document;
        Ok(output)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn seed(count: usize, sections: bool) -> Value {
        let content: Vec<_> = (0..count).map(|index| json!({"type":"paragraph","paraId":format!("{:08X}",index+1),
            "content":[{"type":"run","formatting":{"bold":true},"content":[{"type":"text","text":"Retained authored text. ".repeat(16)}]}]})).collect();
        let mut document = json!({"package":{"document":{"content":content,"retainedBody":true},"retainedPackage":{"opaque":[1,2,3]}},"retainedRoot":"keep"});
        if sections {
            document["package"]["document"]["sections"] = json!([{"content":document["package"]["document"]["content"],"retainedSection":"keep"}]);
        }
        document
    }

    fn reconstruct(before: &Value, response: &Value) -> Value {
        let mut after = before.clone();
        let old = main_content(before).unwrap();
        let content: Vec<_> = response["projectionPatch"]["content"]
            .as_array()
            .unwrap()
            .iter()
            .map(|entry| {
                entry
                    .as_u64()
                    .map(|index| old[index as usize].clone())
                    .unwrap_or_else(|| entry.clone())
            })
            .collect();
        after["package"]["document"]["content"] = json!(content);
        if let Some(metadata) = response["projectionPatch"].get("sectionMetadata") {
            let mut groups = Vec::new();
            let mut current = Vec::new();
            for block in &content {
                current.push(block.clone());
                if block["type"] == "paragraph" && block.get("sectionProperties").is_some() {
                    groups.push(std::mem::take(&mut current));
                }
            }
            if !current.is_empty() || groups.is_empty() {
                groups.push(current);
            }
            let sections: Vec<_> = metadata
                .as_array()
                .unwrap()
                .iter()
                .zip(groups)
                .map(|(section, content)| {
                    let mut section = section.clone();
                    section["content"] = json!(content);
                    section
                })
                .collect();
            after["package"]["document"]["sections"] = json!(sections);
        }
        after
    }

    fn apply_and_compare(
        resident: &mut ResidentDocument,
        before: &Value,
        ops: &[Value],
    ) -> (Value, Value) {
        let expected = apply_operations(before, ops).unwrap();
        let response: Value = serde_json::from_str(
            &resident
                .apply_ops_json(&serde_json::to_string(ops).unwrap())
                .unwrap(),
        )
        .unwrap();
        assert!(response.get("document").is_none());
        assert_eq!(response["inverse"], json!(expected.inverse));
        assert_eq!(
            response["touched"],
            serde_json::to_value(expected.touched).unwrap()
        );
        assert_eq!(response["revisions"], json!(expected.revisions));
        assert_eq!(reconstruct(before, &response), expected.document);
        assert_eq!(
            serde_json::from_str::<Value>(&resident.save().unwrap()).unwrap(),
            expected.document
        );
        (expected.document, response)
    }

    #[test]
    fn forward_reconstruct_undo_and_redo_equal_full_application() {
        for sections in [false, true] {
            let original = seed(3, sections);
            let mut resident = ResidentDocument::load(&original.to_string()).unwrap();
            let operations = [
                json!({"type":"insertText","at":{"story":"main","blockId":"00000002","offset":1},"text":"é😀","runProps":"inherit"}),
                json!({"type":"setParagraphProps","story":"main","blockId":"00000002","patch":{"alignment":"center"}}),
            ];
            let (changed, forward) = apply_and_compare(&mut resident, &original, &operations);
            let (restored, undo) = apply_and_compare(
                &mut resident,
                &changed,
                forward["inverse"].as_array().unwrap(),
            );
            assert_eq!(restored, original);
            let (redone, _) = apply_and_compare(
                &mut resident,
                &restored,
                undo["inverse"].as_array().unwrap(),
            );
            assert_eq!(redone, changed);
        }
    }

    #[test]
    fn atomic_refusal_and_transport_error_preserve_resident_state() {
        let original = seed(2, false);
        let mut resident = ResidentDocument::load(&original.to_string()).unwrap();
        let before = resident.save().unwrap();
        let operations = json!([
            {"type":"insertText","at":{"story":"main","blockId":"00000001","offset":1},"text":"x","runProps":"inherit"},
            {"type":"setParagraphProps","story":"main","blockId":"70000000","patch":{}}
        ]);
        let refused: Value =
            serde_json::from_str(&resident.apply_ops_json(&operations.to_string()).unwrap())
                .unwrap();
        assert_eq!(refused["status"], "refused");
        assert_eq!(resident.save().unwrap(), before);
        assert!(resident.apply_ops_json("not JSON").is_err());
        assert_eq!(resident.save().unwrap(), before);
        let unsupported: Value = serde_json::from_str(
            &resident
                .apply_ops_json("[{\"type\":\"deleteComment\",\"id\":1,\"scope\":\"thread\"}]")
                .unwrap(),
        )
        .unwrap();
        assert_eq!(unsupported["status"], "unsupported");
        assert_eq!(unsupported["dimension"], "residentPackageProjection");
        assert_eq!(resident.save().unwrap(), before);
    }

    #[test]
    fn one_edit_among_a_hundred_blocks_transfers_only_the_changed_block() {
        let original = seed(100, false);
        let mut resident = ResidentDocument::load(&original.to_string()).unwrap();
        let op = json!({"type":"setParagraphProps","story":"main","blockId":"00000033","patch":{"alignment":"center"}});
        let response = resident.apply_ops_json(&json!([op]).to_string()).unwrap();
        let output: Value = serde_json::from_str(&response).unwrap();
        assert!(output.get("document").is_none());
        let content = output["projectionPatch"]["content"].as_array().unwrap();
        assert_eq!(content.iter().filter(|entry| entry.is_object()).count(), 1);
        assert_eq!(content.iter().filter(|entry| entry.is_number()).count(), 99);
        assert!(response.len() < original.to_string().len() / 10);
        assert_eq!(
            reconstruct(&original, &output),
            serde_json::from_str::<Value>(&resident.save().unwrap()).unwrap()
        );
    }

    #[test]
    fn equality_references_survive_top_level_paragraph_movement() {
        let before = seed(3, false);
        let mut after = before.clone();
        after["package"]["document"]["content"] = json!([
            before["package"]["document"]["content"][2],
            before["package"]["document"]["content"][0]
        ]);
        let patch = projection_patch(&before, &after, &json!({"type":"deleteBlocks"})).unwrap();
        assert_eq!(patch.content, json!([2, 0]).as_array().unwrap().clone());
        let response = json!({"projectionPatch":patch});
        assert_eq!(reconstruct(&before, &response), after);
    }
}
