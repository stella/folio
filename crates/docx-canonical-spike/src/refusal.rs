use serde::Serialize;
use serde_json::Value;

/// A TS semantic refusal is separate from a not-yet-implemented spike dimension.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "status", rename_all = "camelCase")]
pub enum Failure {
    Refused {
        #[serde(rename = "opType", skip_serializing_if = "Option::is_none")]
        op_type: Option<String>,
        reason: String,
        message: String,
    },
    Unsupported {
        #[serde(rename = "opType", skip_serializing_if = "Option::is_none")]
        op_type: Option<String>,
        dimension: String,
    },
}

impl Failure {
    pub fn with_op(self, op: &Value) -> Self {
        let op_type = op.get("type").and_then(Value::as_str).map(str::to_owned);
        match self {
            Self::Refused {
                reason, message, ..
            } => Self::Refused {
                op_type,
                reason,
                message,
            },
            Self::Unsupported { dimension, .. } => Self::Unsupported { op_type, dimension },
        }
    }

    pub fn refused(op: &Value, reason: &str, message: impl Into<String>) -> Self {
        Self::Refused {
            op_type: op.get("type").and_then(Value::as_str).map(str::to_owned),
            reason: reason.to_owned(),
            message: message.into(),
        }
    }

    pub fn unsupported(op: &Value, dimension: &str) -> Self {
        Self::Unsupported {
            op_type: op.get("type").and_then(Value::as_str).map(str::to_owned),
            dimension: dimension.to_owned(),
        }
    }
}
