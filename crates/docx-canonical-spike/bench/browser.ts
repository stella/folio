import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";

// Built by the preparation command beside wasm-bindgen's browser-native output.
const loadWasm = async () => {
  const module = await import("/wasm/docx_canonical_spike.js");
  const exports = await module.default();
  return { module, exports };
};

const wasm = await loadWasm();

type SampleRequest = {
  arm: "typescript" | "wasm";
  documentJson: string;
  opsJson: string;
  task: "apply" | "load" | "save";
  verifyOutput?: true;
};

const run = ({ arm, documentJson, opsJson, task, verifyOutput }: SampleRequest) => {
  // A save starts with a retained model. Creating it is outside the timed region.
  const retain = () => {
    if (task !== "save") return undefined;
    if (arm === "wasm") return new wasm.module.JsonModel(documentJson);
    return JSON.parse(documentJson);
  };
  const retained = retain();
  const started = performance.now();
  let output: string;
  if (task === "load") {
    const model =
      arm === "wasm" ? new wasm.module.JsonModel(documentJson) : JSON.parse(documentJson);
    const elapsedMs = performance.now() - started;
    const verifyModel = () => {
      if (!verifyOutput) return undefined;
      if (arm === "wasm") return model.save();
      return JSON.stringify(model);
    };
    const verifiedOutput = verifyModel();
    if (arm === "wasm") model.free();
    return {
      ...(verifyOutput ? { verifiedOutput } : {}),
      elapsedMs,
      outputBytes: new TextEncoder().encode(documentJson).byteLength,
      wasmMemoryBytes: wasm.exports.memory.buffer.byteLength,
    };
  }
  if (task === "save") {
    output = arm === "wasm" ? retained.save() : JSON.stringify(retained);
    const elapsedMs = performance.now() - started;
    if (arm === "wasm") retained.free();
    return {
      ...(verifyOutput ? { verifiedOutput: output } : {}),
      elapsedMs,
      outputBytes: new TextEncoder().encode(output).byteLength,
      wasmMemoryBytes: wasm.exports.memory.buffer.byteLength,
    };
  }
  switch (arm) {
    case "typescript": {
      const document = JSON.parse(documentJson);
      const ops = JSON.parse(opsJson);
      const result = applyDocumentOps(document, ops);
      if (result.isErr()) return { status: "refused", reason: result.error.reason };
      output = JSON.stringify(result.value);
      break;
    }
    case "wasm":
      output = wasm.module.apply(documentJson, opsJson);
      break;
  }
  const elapsedMs = performance.now() - started;
  const result: unknown = JSON.parse(output);
  if (typeof result === "object" && result !== null && "status" in result) return result;
  return {
    elapsedMs,
    outputBytes: new TextEncoder().encode(output).byteLength,
    wasmMemoryBytes: wasm.exports.memory.buffer.byteLength,
  };
};

Object.defineProperty(globalThis, "runCanonicalSpikeSample", { value: run });
Object.defineProperty(globalThis, "canonicalSpikeReady", { value: true });

const canonical = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(canonical);
  if (typeof value === "object" && value !== null)
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, child]) => [key, canonical(child)]),
    );
  return value;
};
const verify = ({
  arm,
  documentJson,
  opsJson,
  expectedJson,
}: {
  arm: "typescript" | "wasm";
  documentJson: string;
  opsJson: string;
  expectedJson: string;
}) => {
  const expected = JSON.stringify(canonical(JSON.parse(expectedJson)));
  let output;
  if (arm === "wasm") output = wasm.module.apply(documentJson, opsJson);
  else {
    const applied = applyDocumentOps(JSON.parse(documentJson), JSON.parse(opsJson));
    if (applied.isErr()) return false;
    output = JSON.stringify(applied.value);
  }
  return JSON.stringify(canonical(JSON.parse(output))) === expected;
};
Object.defineProperty(globalThis, "verifyCanonicalSpikeArm", { value: verify });
Object.defineProperty(globalThis, "verifyCanonicalSpikeReplay", {
  value: ({
    arm,
    documentJson,
    task,
  }: {
    arm: "typescript" | "wasm";
    documentJson: string;
    task: "load" | "save";
  }) => {
    const result = run({ arm, documentJson, opsJson: "[]", task, verifyOutput: true });
    if (!("verifiedOutput" in result) || typeof result.verifiedOutput !== "string") return false;
    return (
      JSON.stringify(canonical(JSON.parse(result.verifiedOutput))) ===
      JSON.stringify(canonical(JSON.parse(documentJson)))
    );
  },
});
