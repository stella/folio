import { projectionPatch, sameValue } from "./resident-projection";
import type { Document } from "../../../packages/docx-core/src/model/document";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";

// Built by the preparation command beside wasm-bindgen's browser-native output.
const loadWasm = async () => {
  const module = await import("/wasm/docx_canonical_spike.js");
  const exports = await module.default();
  return { module, exports };
};

const wasm = await loadWasm();

let residentWasm: InstanceType<typeof wasm.module.ResidentModel> | undefined;
let residentTs: Document | undefined;
const initializeResidents = (documentJson: string) => {
  residentWasm?.free();
  residentWasm = new wasm.module.ResidentModel(documentJson);
  residentTs = JSON.parse(documentJson);
};
Object.defineProperty(globalThis, "initializeCanonicalSpikeResidents", {
  value: initializeResidents,
});

const residentSample = ({ arm, opsJson }: { arm: "typescript" | "wasm"; opsJson: string }) => {
  if (!residentWasm || !residentTs) throw new TypeError("Resident fixture must be loaded first.");
  const before = residentTs;
  const started = performance.now();
  let output: string;
  if (arm === "wasm") output = residentWasm.apply(opsJson);
  else {
    const result = applyDocumentOps(residentTs, JSON.parse(opsJson));
    if (result.isErr()) throw new TypeError(result.error.message);
    const { document, inverse, touched, revisions } = result.value;
    output = JSON.stringify({
      inverse,
      touched,
      revisions,
      projectionPatch: projectionPatch(residentTs, document),
    });
    residentTs = document;
  }
  // Decoding the inverse/patch is part of the resident boundary cost.
  const result = JSON.parse(output);
  const elapsedMs = performance.now() - started;
  if (result.status !== undefined) throw new TypeError(`Resident apply failed: ${output}`);
  const inverseJson = JSON.stringify(result.inverse);
  // Restore outside the sample so every repetition uses the identical loaded model.
  if (arm === "wasm") {
    const undo = JSON.parse(residentWasm.apply(inverseJson));
    if (undo.status !== undefined) throw new TypeError("Resident undo failed.");
  } else {
    residentTs = applyDocumentOps(residentTs, JSON.parse(inverseJson)).unwrap().document;
    if (!sameValue(residentTs, before))
      throw new TypeError("Resident TS undo did not restore fixture.");
  }
  return {
    elapsedMs,
    outputBytes: new TextEncoder().encode(output).byteLength,
    wasmMemoryBytes: wasm.exports.memory.buffer.byteLength,
  };
};
Object.defineProperty(globalThis, "runCanonicalSpikeResidentSample", { value: residentSample });
Object.defineProperty(globalThis, "verifyCanonicalSpikeResident", {
  value: ({
    documentJson,
    opsJson,
    expectedJson,
  }: {
    documentJson: string;
    opsJson: string;
    expectedJson: string;
  }) => {
    const model = new wasm.module.ResidentModel(documentJson);
    try {
      const before: Document = JSON.parse(documentJson);
      const expected = JSON.parse(expectedJson);
      const actual = JSON.parse(model.apply(opsJson));
      if (
        !sameValue(actual, {
          inverse: expected.inverse,
          touched: expected.touched,
          revisions: expected.revisions,
          projectionPatch: projectionPatch(before, expected.document),
        })
      )
        return false;
      if (!sameValue(JSON.parse(model.save()), expected.document)) return false;
      const patch = actual.projectionPatch.content.map((block: unknown) =>
        typeof block === "number" ? before.package.document.content.at(block) : block,
      );
      if (!sameValue(patch, expected.document.package.document.content)) return false;
      return true;
    } finally {
      model.free();
    }
  },
});

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
  const result: unknown = JSON.parse(output);
  const elapsedMs = performance.now() - started;
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
