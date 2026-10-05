//! Self-timing native arm. Startup/IPC are outside samples, JSON costs are inside.
use serde_json::{Value, json};
use std::time::Instant;

fn load_model(document: &str) -> Result<Value, serde_json::Error> {
    serde_json::from_str(document)
}
fn save_model(model: &Value) -> Result<String, serde_json::Error> {
    serde_json::to_string(model)
}

pub fn run(request: &Value) -> Value {
    let Some(document) = request.get("documentJson").and_then(Value::as_str) else {
        return json!({"status":"transportError","message":"Missing documentJson"});
    };
    let Some(ops) = request.get("opsJson").and_then(Value::as_str) else {
        return json!({"status":"transportError","message":"Missing opsJson"});
    };
    let iterations = request
        .get("iterations")
        .and_then(Value::as_u64)
        .unwrap_or(1)
        .clamp(1, 1000);
    let mut samples = Vec::new();
    let mut output_bytes = 0;
    let task = request
        .get("task")
        .and_then(Value::as_str)
        .unwrap_or("apply");
    if task == "resident" && request.get("verify").and_then(Value::as_bool) == Some(true) {
        let checked = (|| -> Result<Value, String> {
            let mut model = crate::resident::ResidentDocument::load(document)?;
            let result: Value = serde_json::from_str(&model.apply_ops_json(ops)?)
                .map_err(|error| error.to_string())?;
            let document: Value =
                serde_json::from_str(&model.save()?).map_err(|error| error.to_string())?;
            Ok(json!({"residentResult":result,"document":document}))
        })();
        return checked
            .unwrap_or_else(|message| json!({"status":"transportError","message":message}));
    }
    // Untimed model verification exercises the same load/save functions as samples.
    if request.get("verify").and_then(Value::as_bool) == Some(true) {
        let parsed = match load_model(document) {
            Ok(value) => value,
            Err(error) => return json!({"status":"transportError","message":error.to_string()}),
        };
        return match task {
            "load" => json!({"verifiedModel":parsed}),
            "save" => {
                match save_model(&parsed).and_then(|saved| serde_json::from_str::<Value>(&saved)) {
                    Ok(value) => json!({"verifiedModel":value}),
                    Err(error) => json!({"status":"transportError","message":error.to_string()}),
                }
            }
            _ => {
                json!({"status":"transportError","message":"Model verification expects load or save"})
            }
        };
    }
    if task == "resident" {
        let mut model = match crate::resident::ResidentDocument::load(document) {
            Ok(model) => model,
            Err(message) => return json!({"status":"transportError","message":message}),
        };
        for _ in 0..iterations {
            let started = Instant::now();
            let output = match model.apply_ops_json(ops) {
                Ok(output) => output,
                Err(message) => return json!({"status":"transportError","message":message}),
            };
            let result: Value = match serde_json::from_str(&output) {
                Ok(result) => result,
                Err(error) => return json!({"status":"transportError","message":error.to_string()}),
            };
            let elapsed = started.elapsed().as_secs_f64() * 1000.0;
            if result.get("status").is_some() {
                return result;
            }
            samples.push(elapsed);
            output_bytes = output.len();
            let inverse = match serde_json::to_string(&result["inverse"]) {
                Ok(inverse) => inverse,
                Err(error) => return json!({"status":"transportError","message":error.to_string()}),
            };
            let undo = match model.apply_ops_json(&inverse) {
                Ok(undo) => undo,
                Err(message) => return json!({"status":"transportError","message":message}),
            };
            let undo: Value = match serde_json::from_str(&undo) {
                Ok(undo) => undo,
                Err(error) => return json!({"status":"transportError","message":error.to_string()}),
            };
            if undo.get("status").is_some() {
                return undo;
            }
        }
        return json!({"samplesMs":samples,"outputBytes":output_bytes});
    }
    let model = if task == "save" {
        match load_model(document) {
            Ok(value) => Some(value),
            Err(error) => return json!({"status":"transportError","message":error.to_string()}),
        }
    } else {
        None
    };
    for _ in 0..iterations {
        let start = Instant::now();
        let output = match match task {
            "apply" => crate::apply_json(document, ops),
            "load" => {
                let parsed = load_model(document);
                let elapsed_ms = start.elapsed().as_secs_f64() * 1000.0;
                if let Err(error) = parsed {
                    return json!({"status":"transportError","message":error.to_string()});
                }
                samples.push(elapsed_ms);
                output_bytes = document.len();
                continue;
            }
            "save" => save_model(model.as_ref().expect("save model parsed"))
                .map_err(|error| error.to_string()),
            _ => return json!({"status":"transportError","message":"Unknown benchmark task"}),
        } {
            Ok(output) => output,
            Err(message) => return json!({"status":"transportError","message":message}),
        };
        let elapsed_ms = start.elapsed().as_secs_f64() * 1000.0;
        // Never include unsupported/refused work in latency comparisons.
        let check: Value = match serde_json::from_str(&output) {
            Ok(check) => check,
            Err(error) => return json!({"status":"transportError","message":error.to_string()}),
        };
        if task == "apply" && check.get("status").is_some() {
            return check;
        }
        let elapsed_ms = if task == "apply" {
            start.elapsed().as_secs_f64() * 1000.0
        } else {
            elapsed_ms
        };
        samples.push(elapsed_ms);
        output_bytes = output.len();
    }
    json!({"samplesMs":samples,"outputBytes":output_bytes})
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn untimed_load_save_verification_uses_the_measured_model_functions() {
        let document = json!({"package":{"document":{"content":[]}},"facts":[null,1,"é😀"],"owned":{"null":null}});
        for task in ["load", "save"] {
            let checked = run(
                &json!({"documentJson":document.to_string(),"opsJson":"[]","task":task,"verify":true}),
            );
            assert_eq!(checked, json!({"verifiedModel":document}));
        }
    }
}
