use docx_canonical_spike::{apply::apply_operations, refusal::Failure};
use serde_json::{Value, json};

fn document() -> Value {
    json!({"package":{"document":{"content":[
        {"type":"paragraph","paraId":"00000001","content":[{"type":"run","content":[{"type":"text","text":"alpha"}]}],"formatting":{"bold":true},"unmodeled":{"owner":"keep"}},
        {"type":"paragraph","paraId":"00000002","content":[]}
    ]},"relationships":[{"id":"rId1","target":"keep.xml"}]},"warnings":["preserved"]})
}

#[test]
fn inverse_and_inverse_of_inverse_preserve_every_unrelated_field() {
    for base in [
        None,
        Some(json!({})),
        Some(json!({"alignment":"start","keepNext":false})),
    ] {
        for patch in [
            json!({}),
            json!({"alignment":"center"}),
            json!({"alignment":null}),
            json!({"keepNext":true,"alignment":null}),
        ] {
            let mut before = document();
            if let Some(base) = base.clone() {
                before["package"]["document"]["content"][0]["formatting"] = base;
            } else {
                before["package"]["document"]["content"][0]
                    .as_object_mut()
                    .expect("paragraph")
                    .remove("formatting");
            }
            let op = json!({"type":"setParagraphProps","story":"main","blockId":"00000001","patch":patch});
            let applied = apply_operations(&before, &[op]).expect("apply");
            let restored = apply_operations(&applied.document, &applied.inverse).expect("inverse");
            assert_eq!(restored.document, before);
            let redone = apply_operations(&restored.document, &restored.inverse).expect("redo");
            assert_eq!(redone.document, applied.document);
            assert_eq!(
                applied.document["package"]["relationships"],
                before["package"]["relationships"]
            );
            assert_eq!(
                applied.document["package"]["document"]["content"][1],
                before["package"]["document"]["content"][1]
            );
        }
    }
}

#[test]
fn batch_refusal_does_not_mutate_the_input() {
    let before = document();
    let original = before.clone();
    let ops = [
        json!({"type":"setParagraphProps","story":"main","blockId":"00000001","patch":{"alignment":"center"}}),
        json!({"type":"setParagraphProps","story":"main","blockId":"00000003","patch":{}}),
    ];
    assert!(
        matches!(apply_operations(&before,&ops),Err(Failure::Refused{reason,..}) if reason=="blockNotFound")
    );
    assert_eq!(before, original);
}

#[test]
fn stale_precondition_refuses_without_touching_the_document() {
    let before = document();
    let op = json!({"type":"setParagraphProps","story":"main","blockId":"00000001","expected":{"bold":false},"patch":{"alignment":"center"}});
    assert!(
        matches!(apply_operations(&before,&[op]),Err(Failure::Refused{reason,..}) if reason=="stale")
    );
}

#[test]
fn unsupported_dimension_is_not_a_semantic_refusal() {
    let before = document();
    let op = json!({"type":"notImplementedSpikeDimension","at":{"story":"main","blockId":"00000001","offset":0},"text":"x","runProps":"inherit"});
    assert!(matches!(
        apply_operations(&before, &[op]),
        Err(Failure::Unsupported { .. })
    ));
}
