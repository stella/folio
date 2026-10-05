import { test, expect } from "bun:test";
import { deepStrictEqual, throws } from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import fc from "fast-check";
import {
  documentArbitrary,
  opSeedArbitrary,
  opFor,
} from "../../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";
import {
  packageDocumentArbitrary,
  captureDocumentArbitrary,
} from "../../../test/generators/packageOperationArbitraries";
import {
  HarnessCodecError,
  decodeTagged,
  encodeTagged,
  withoutCaptureSymbols,
  transportScope,
} from "./harness-codec";

const parseNativeResponses = (output: string, expectedCount: number): unknown[] => {
  const body = output.trimEnd();
  const lines = body.length === 0 ? [] : body.split("\n");
  expect(lines, "Native JSONL must return exactly one response per request.").toHaveLength(
    expectedCount,
  );
  return lines.map((line): unknown => JSON.parse(line));
};

const nativeResponses = (encoded: readonly unknown[]): unknown[] => {
  if (encoded.length === 0) return [];
  const input = `${encoded.map((roundtrip) => JSON.stringify({ harness: { roundtrip } })).join("\n")}\n`;
  const result = spawnSync(
    process.env["RUST_SPIKE_NATIVE_BINARY"] ??
      fileURLToPath(new URL("../target/debug/canonical-spike", import.meta.url)),
    [],
    {
      input,
      encoding: "utf8",
      // Roundtrip output is smaller than its request; leave room for error replies.
      maxBuffer: Math.max(1024 * 1024, Buffer.byteLength(input) + encoded.length * 1024),
    },
  );
  if (result.error) throw result.error;
  expect(result.status, result.stderr).toBe(0);
  return parseNativeResponses(result.stdout, encoded.length);
};

const nativeRoundtrips = (values: readonly unknown[]): unknown[] =>
  nativeResponses(values.map(encodeTagged)).map((response, index) => {
    if (typeof response !== "object" || response === null || !("harness" in response))
      throw new TypeError(`Rust codec did not roundtrip model ${index}.`);
    return decodeTagged(response.harness);
  });
const nativeRoundtrip = (value: unknown): unknown => nativeRoundtrips([value]).at(0);

test("native JSONL response count and positional order remain exact", () => {
  deepStrictEqual(parseNativeResponses('1\n{"harness":null}\n', 2), [1, { harness: null }]);
  throws(() => parseNativeResponses("1\n", 2));
  throws(() => parseNativeResponses("1\n2\n3\n", 2));
  throws(() => parseNativeResponses("1\n\n2\n", 3), SyntaxError);
});

for (const [family, arbitrary] of [
  ["document", documentArbitrary],
  ["package", packageDocumentArbitrary],
  ["capture", captureDocumentArbitrary],
] as const) {
  test(`tagged harness roundtrips existing ${family} arbitrary with exact own-field presence`, () => {
    const documents = [20261005, 20261006, 327444275, -392419793].flatMap((seed) =>
      fc.sample(arbitrary, { seed, numRuns: 50 }),
    );
    expect(documents).toHaveLength(200);
    const native = nativeRoundtrips(documents);
    for (const [index, document] of documents.entries()) {
      deepStrictEqual(
        decodeTagged(JSON.parse(JSON.stringify(encodeTagged(document)))),
        withoutCaptureSymbols(document),
      );
      deepStrictEqual(native.at(index), withoutCaptureSymbols(document));
      deepStrictEqual(
        encodeTagged(native.at(index)),
        encodeTagged(withoutCaptureSymbols(document)),
      );
    }
  });
}
const nullPrototypeCount = (value: unknown): number => {
  if (typeof value !== "object" || value === null) return 0;
  if (value instanceof Date || value instanceof Uint8Array || value instanceof ArrayBuffer)
    return 0;
  if (value instanceof Map)
    return [...value.values()].reduce((count, entry) => count + nullPrototypeCount(entry), 0);
  return Object.values(value).reduce(
    (count, entry) => count + nullPrototypeCount(entry),
    Object.getPrototypeOf(value) === null ? 1 : 0,
  );
};

test("TS codec preserves null prototypes in generated edit outputs, including CI seed 20261005", () => {
  let nullObjects = 0;
  for (const seed of [20261005, 20261006, 327444275, -392419793]) {
    const draws = fc.sample(fc.tuple(documentArbitrary, opSeedArbitrary), { seed, numRuns: 80 });
    for (const [document, generated] of draws) {
      const result = applyDocumentOps(document, [opFor(document, generated)]);
      if (result.isErr()) continue;
      nullObjects += nullPrototypeCount(result.value.document);
      const encoded = encodeTagged(result.value.document);
      const decoded = decodeTagged(JSON.parse(JSON.stringify(encoded)));
      deepStrictEqual(decoded, withoutCaptureSymbols(result.value.document));
      deepStrictEqual(encodeTagged(decoded), encoded);
    }
  }
  expect(nullObjects).toBeGreaterThan(0);
});

test("Rust codec preserves prototypes across generated edit outputs", () => {
  const documents = [];
  for (const seed of [20261005, 20261006, 327444275, -392419793])
    for (const [document, generated] of fc.sample(fc.tuple(documentArbitrary, opSeedArbitrary), {
      seed,
      numRuns: 80,
    })) {
      const result = applyDocumentOps(document, [opFor(document, generated)]);
      if (result.isErr()) continue;
      documents.push(result.value.document);
    }
  expect(documents.length).toBeGreaterThan(0);
  const native = nativeRoundtrips(documents);
  for (const [index, document] of documents.entries()) {
    deepStrictEqual(native.at(index), withoutCaptureSymbols(document));
    deepStrictEqual(encodeTagged(native.at(index)), encodeTagged(document));
  }
});

test("null-prototype records retain prototype, special keys and owned undefined", () => {
  const record = Object.setPrototypeOf(
    Object.fromEntries([
      ["__proto__", { tag: "object", entries: "authored" }],
      ["constructor", undefined],
      ["toString", null],
    ]),
    null,
  );
  const original = { record, nested: new Map([["entry", [record]]]) };
  const encoded = encodeTagged(original);
  for (const decoded of [decodeTagged(encoded), nativeRoundtrip(original)]) {
    deepStrictEqual(decoded, original);
    deepStrictEqual(decoded, withoutCaptureSymbols(original));
    deepStrictEqual(encodeTagged(decoded), encoded);
  }
  expect(encodeTagged(record)).not.toEqual(encodeTagged({ ...record }));
  expect(Object.getPrototypeOf(withoutCaptureSymbols(record))).toBeNull();
});

test("unsupported codec values report type and prototype without invoking accessors", () => {
  const prototype = {};
  Object.defineProperty(prototype, "constructor", {
    get: () => {
      throw new HarnessCodecError({ message: "Constructor accessor must not run." });
    },
  });
  const value = Object.setPrototypeOf({}, prototype);
  throws(
    () => encodeTagged(value),
    (error: unknown) =>
      error instanceof HarnessCodecError &&
      error.message === "Unsupported model value in harness codec (type object, prototype custom).",
  );
});

test("tag-shaped authored records, Map order, binary, Date and undefined never collapse", () => {
  const mapValue = (value: null | undefined) => ({ value });
  const sparse: unknown[] = [];
  sparse.length = 3;
  sparse[1] = undefined;
  sparse[2] = null;
  const original = {
    tag: "map",
    entries: "authored",
    missing: undefined,
    nil: null,
    date: new Date("2026-01-02T03:04:05Z"),
    map: new Map([
      ["z", mapValue(undefined)],
      ["a", mapValue(null)],
    ]),
    bytes: new Uint8Array([0, 255]),
    buffer: new Uint8Array([3, 2, 1]).buffer,
    sparse,
  };
  deepStrictEqual(decodeTagged(JSON.parse(JSON.stringify(encodeTagged(original)))), original);
  deepStrictEqual(nativeRoundtrip(original), original);
  expect(encodeTagged({ field: undefined })).not.toEqual(encodeTagged({}));
});
test("Rust preserves Map entry order with undefined entries in every position", () => {
  const keys = ["z", "a", "m"];
  const originals = [];
  for (let mask = 0; mask < 8; mask += 1) {
    const entries = keys.map(
      (key, index) =>
        [key, mask & (1 << index) ? undefined : index] satisfies [string, number | undefined],
    );
    originals.push(new Map(entries));
  }
  const native = nativeRoundtrips(originals);
  for (const [index, original] of originals.entries())
    deepStrictEqual(encodeTagged(native.at(index)), encodeTagged(original));
});
test("capture-symbol and shared-reference exclusions are counted", () => {
  const shared = { x: 1 };
  const original = { a: shared, b: shared, [Symbol("capture")]: "opaque" };
  expect(transportScope(original)).toEqual({
    excludedCaptureSymbols: 1,
    sharedReferences: 1,
    nonEnumerableStringFields: 0,
    accessorFields: 0,
  });
  deepStrictEqual(decodeTagged(encodeTagged(original)), { a: { x: 1 }, b: { x: 1 } });
});
test("nonenumerable owned undefined data remains an owned field and is counted", () => {
  const child = { visible: true };
  Object.defineProperty(child, "hidden", { value: undefined, enumerable: false });
  const original = { visible: true };
  Object.defineProperty(original, "hidden", { value: undefined, enumerable: false });
  Object.defineProperty(original, "child", { value: child, enumerable: false });
  expect(transportScope(original)).toEqual({
    excludedCaptureSymbols: 0,
    sharedReferences: 0,
    nonEnumerableStringFields: 3,
    accessorFields: 0,
  });
  const encoded = encodeTagged(original);
  for (const decoded of [decodeTagged(encoded), nativeRoundtrip(original)]) {
    deepStrictEqual(decoded, withoutCaptureSymbols(original));
    deepStrictEqual(encodeTagged(decoded), encoded);
    if (typeof decoded !== "object" || decoded === null)
      throw new HarnessCodecError({ message: "Expected decoded object." });
    expect(Object.hasOwn(decoded, "hidden")).toBe(true);
    expect(Reflect.get(decoded, "hidden")).toBeUndefined();
  }
});

const MALFORMED_CORPUS = [
  {
    name: "null object missing entries",
    value: { tag: "nullObject" },
    message: "Unexpected fields in the nullObject tag.",
  },
  {
    name: "null object extra field",
    value: { tag: "nullObject", entries: [], prototype: null },
    message: "Unexpected fields in the nullObject tag.",
  },
  {
    name: "duplicate null object key",
    value: {
      tag: "nullObject",
      entries: [
        ["x", 1],
        ["x", 2],
      ],
    },
    message: "Harness entries repeat an owned key.",
  },
  { name: "raw array", value: [], message: "Raw arrays are not tagged harness values." },
  { name: "missing discriminator", value: {}, message: "A harness object needs a tag." },
  { name: "nonstring discriminator", value: { tag: 3 }, message: "A harness object needs a tag." },
  {
    name: "unknown discriminator",
    value: { tag: "unknown" },
    message: "Unknown harness tag unknown.",
  },
  {
    name: "undefined extra field",
    value: { tag: "undefined", extra: true },
    message: "Unexpected fields in the undefined tag.",
  },
  {
    name: "missing date field",
    value: { tag: "date" },
    message: "Unexpected fields in the date tag.",
  },
  {
    name: "date extra field",
    value: { tag: "date", iso: "2026-01-01T00:00:00.000Z", extra: true },
    message: "Unexpected fields in the date tag.",
  },
  { name: "nonstring date", value: { tag: "date", iso: 3 }, message: "Date ISO must be a string." },
  { name: "invalid date", value: { tag: "date", iso: "invalid" }, message: "Invalid ISO Date." },
  {
    name: "noncanonical date",
    value: { tag: "date", iso: "2026-01-01T00:00:00Z" },
    message: "Invalid ISO Date.",
  },
  {
    name: "calendar overflow",
    value: { tag: "date", iso: "1900-02-29T00:00:00.000Z" },
    message: "Invalid ISO Date.",
  },
  {
    name: "date range overflow",
    value: { tag: "date", iso: "+275760-09-13T00:00:00.001Z" },
    message: "Invalid ISO Date.",
  },
  {
    name: "binary nonarray",
    value: { tag: "uint8Array", bytes: null },
    message: "Binary bytes must be an array.",
  },
  {
    name: "binary overflow",
    value: { tag: "uint8Array", bytes: [256] },
    message: "Binary bytes must be integers in 0..255.",
  },
  {
    name: "buffer noninteger",
    value: { tag: "arrayBuffer", bytes: [0.5] },
    message: "Binary bytes must be integers in 0..255.",
  },
  {
    name: "array nonarray items",
    value: { tag: "array", items: null },
    message: "Array items must be an array.",
  },
  {
    name: "root hole",
    value: { tag: "hole" },
    message: "A hole tag belongs only to an array slot.",
  },
  {
    name: "array hole extra field",
    value: { tag: "array", items: [{ tag: "hole", extra: true }] },
    message: "Unexpected fields in the hole tag.",
  },
  {
    name: "object hole",
    value: { tag: "object", entries: [["x", { tag: "hole" }]] },
    message: "A hole tag belongs only to an array slot.",
  },
  {
    name: "map nonarray entries",
    value: { tag: "map", entries: null },
    message: "Object and Map entries must be arrays.",
  },
  {
    name: "nonpair entry",
    value: { tag: "object", entries: [["x"]] },
    message: "An entry must be a key/value pair.",
  },
  {
    name: "nonstring key",
    value: { tag: "map", entries: [[3, 1]] },
    message: "Harness object and Map keys must be strings.",
  },
  {
    name: "duplicate map key",
    value: {
      tag: "map",
      entries: [
        ["x", 1],
        ["x", 2],
      ],
    },
    message: "Harness entries repeat an owned key.",
  },
  {
    name: "duplicate object key",
    value: {
      tag: "object",
      entries: [
        ["x", 1],
        ["x", 2],
      ],
    },
    message: "Harness entries repeat an owned key.",
  },
] as const;

test("the shared malformed corpus fails with exact errors in TS and Rust", () => {
  const native = nativeResponses(MALFORMED_CORPUS.map(({ value }) => value));
  for (const [index, { name, value, message }] of MALFORMED_CORPUS.entries()) {
    throws(
      () => decodeTagged(value),
      (error: unknown) => error instanceof HarnessCodecError && error.message === message,
      name,
    );
    deepStrictEqual(native.at(index), { status: "transportError", message }, name);
  }
});
