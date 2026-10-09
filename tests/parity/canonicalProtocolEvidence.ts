import { stripVTControlCharacters } from "node:util";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

const CONTEXT_EVENTS = new Set([
  "Runtime.executionContextCreated",
  "Runtime.executionContextDestroyed",
  "Runtime.executionContextsCleared",
  "Page.frameNavigated",
  "Page.frameStartedLoading",
  "Page.frameStoppedLoading",
  "Inspector.targetCrashed",
  "Target.detachedFromTarget",
]);

const evaluationOperation = (expressions: string[]) => {
  for (const [operation, call] of [
    ["canonical-load", ".load("],
    ["canonical-save", ".save("],
    ["canonical-snapshot", ".snapshot("],
  ] as const) {
    if (expressions.some((value) => value.includes("__folioCanonical") && value.includes(call)))
      return operation;
  }
  return "other";
};

/** Preserve protocol errors and their calls without transporting document or screenshot payloads. */
export const canonicalProtocolEvidence = (line: string) => {
  const text = stripVTControlCharacters(line);
  const marker = text.indexOf("pw:protocol");
  if (marker === -1) return { type: "output", text: line } as const;
  const start = text.indexOf("{", marker);
  const end = text.lastIndexOf("}");
  if (start === -1 || end < start) return { type: "unparsed", length: text.length } as const;
  const parsed: unknown = JSON.parse(text.slice(start, end + 1));
  if (!isRecord(parsed)) return { type: "unparsed", length: text.length } as const;
  const base = { id: parsed["id"], sessionId: parsed["sessionId"] };
  if (parsed["error"] !== undefined)
    return { type: "evidence", message: { ...base, error: parsed["error"] } } as const;
  const method = parsed["method"];
  const params = parsed["params"];
  if (typeof method !== "string") return { type: "discard" } as const;
  if (CONTEXT_EVENTS.has(method))
    return { type: "evidence", message: { ...base, method, params } } as const;
  if (method !== "Runtime.callFunctionOn" && method !== "Runtime.evaluate")
    return { type: "discard" } as const;
  if (!isRecord(params)) return { type: "unparsed", length: text.length } as const;
  const args = params["arguments"];
  const expressions = Array.isArray(args)
    ? args.flatMap((arg: unknown) =>
        isRecord(arg) && typeof arg["value"] === "string" ? [arg["value"]] : [],
      )
    : [];
  const expression = params["expression"];
  if (typeof expression === "string") expressions.push(expression);
  return {
    type: "evidence",
    message: {
      ...base,
      method,
      contextId: params["contextId"],
      objectId: params["objectId"],
      awaitPromise: params["awaitPromise"],
      operation: evaluationOperation(expressions),
    },
  } as const;
};
