//! Canonical-core experiment. No production package imports this crate.
pub mod apply;
pub mod benchmark;
mod blocks;
mod comments;
pub mod harness;
mod identity;
pub mod inline;
pub mod model;
mod paragraphs;
pub mod refusal;
mod run_props;
mod tables;
mod tracked;
pub mod wire;

use serde_json::{Value, json};

/// Both native and browser entry points use this exact JSON boundary.
pub fn apply_json(document_json: &str, operations_json: &str) -> Result<String, String> {
    let document: Value = serde_json::from_str(document_json).map_err(|error| error.to_string())?;
    let operations: Vec<Value> =
        serde_json::from_str(operations_json).map_err(|error| error.to_string())?;
    let result = match apply::apply_operations(&document, &operations) {
        Ok(applied) => serde_json::to_value(applied).map_err(|error| error.to_string())?,
        Err(failure) => serde_json::to_value(failure).map_err(|error| error.to_string())?,
    };
    serde_json::to_string(&result).map_err(|error| error.to_string())
}

/// JSON-lines transport for differential cases; transport errors never impersonate op refusals.
pub fn apply_request(request: &Value) -> Value {
    match (
        request.get("document"),
        request.get("ops").and_then(Value::as_array),
    ) {
        (Some(document), Some(ops)) => match apply::apply_operations(document, ops) {
            Ok(applied) => serde_json::to_value(applied).unwrap_or_else(
                |error| json!({"status":"transportError","message":error.to_string()}),
            ),
            Err(error) => serde_json::to_value(error).unwrap_or_else(
                |error| json!({"status":"transportError","message":error.to_string()}),
            ),
        },
        _ => json!({"status":"transportError","message":"Expected document and ops array."}),
    }
}

#[cfg(all(feature = "wasm", target_arch = "wasm32"))]
#[wasm_bindgen::prelude::wasm_bindgen]
pub fn apply(document_json: &str, operations_json: &str) -> Result<String, wasm_bindgen::JsValue> {
    apply_json(document_json, operations_json)
        .map_err(|error| wasm_bindgen::JsValue::from_str(&error))
}

/// Harness-only tagged transport; never part of a published package contract.
#[cfg(all(feature = "wasm", target_arch = "wasm32"))]
#[wasm_bindgen::prelude::wasm_bindgen]
pub fn apply_harness_json(request_json: &str) -> Result<String, wasm_bindgen::JsValue> {
    let request: Value = serde_json::from_str(request_json)
        .map_err(|error| wasm_bindgen::JsValue::from_str(&error.to_string()))?;
    serde_json::to_string(&harness::apply_request(&request))
        .map_err(|error| wasm_bindgen::JsValue::from_str(&error.to_string()))
}

/// Retained lossless JSON model, so load and replay-save can be measured separately.
/// This deliberately measures model transfer, not ZIP/XML parsing or DOCX serialization.
#[cfg(all(feature = "wasm", target_arch = "wasm32"))]
#[wasm_bindgen::prelude::wasm_bindgen]
pub struct JsonModel {
    value: Value,
}

#[cfg(all(feature = "wasm", target_arch = "wasm32"))]
#[wasm_bindgen::prelude::wasm_bindgen]
impl JsonModel {
    #[wasm_bindgen::prelude::wasm_bindgen(constructor)]
    pub fn new(document_json: &str) -> Result<JsonModel, wasm_bindgen::JsValue> {
        let value = serde_json::from_str(document_json)
            .map_err(|error| wasm_bindgen::JsValue::from_str(&error.to_string()))?;
        Ok(Self { value })
    }

    pub fn save(&self) -> Result<String, wasm_bindgen::JsValue> {
        serde_json::to_string(&self.value)
            .map_err(|error| wasm_bindgen::JsValue::from_str(&error.to_string()))
    }
}
