import { test, expect } from "bun:test";
import { deepStrictEqual, throws } from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import fc from "fast-check";
import { documentArbitrary } from "../../../packages/docx-core/src/ops/__tests__/documentArbitraries";
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

const nativeResponse = (encoded: unknown): unknown => {
  const result = spawnSync(
    process.env["RUST_SPIKE_NATIVE_BINARY"] ??
      fileURLToPath(new URL("../target/debug/canonical-spike", import.meta.url)),
    [],
    { input: `${JSON.stringify({ harness: { roundtrip: encoded } })}\n`, encoding: "utf8" },
  );
  if (result.error) throw result.error;
  expect(result.status, result.stderr).toBe(0);
  const response: unknown = JSON.parse(result.stdout);
  return response;
};

const nativeRoundtrip = (value: unknown): unknown => {
  const response = nativeResponse(encodeTagged(value));
  if (typeof response !== "object" || response === null || !("harness" in response))
    throw new TypeError("Rust codec did not roundtrip the model.");
  return decodeTagged(response.harness);
};

for (const [family, arbitrary] of [
  ["document", documentArbitrary],
  ["package", packageDocumentArbitrary],
  ["capture", captureDocumentArbitrary],
] as const) {
  test(`tagged harness roundtrips existing ${family} arbitrary with exact own-field presence`, () => {
    for (const seed of [20261005, 20261006, 327444275, -392419793])
      for (const document of fc.sample(arbitrary, { seed, numRuns: 50 })) {
        deepStrictEqual(
          decodeTagged(JSON.parse(JSON.stringify(encodeTagged(document)))),
          withoutCaptureSymbols(document),
        );
        const native = nativeRoundtrip(document);
        deepStrictEqual(native, withoutCaptureSymbols(document));
        deepStrictEqual(encodeTagged(native), encodeTagged(withoutCaptureSymbols(document)));
      }
  });
}
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
  for (let mask = 0; mask < 8; mask += 1) {
    const entries = keys.map(
      (key, index) =>
        [key, mask & (1 << index) ? undefined : index] satisfies [string, number | undefined],
    );
    const original = new Map(entries);
    deepStrictEqual(encodeTagged(nativeRoundtrip(original)), encodeTagged(original));
  }
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
  for (const { name, value, message } of MALFORMED_CORPUS) {
    throws(
      () => decodeTagged(value),
      (error: unknown) => error instanceof HarnessCodecError && error.message === message,
      name,
    );
    deepStrictEqual(nativeResponse(value), { status: "transportError", message }, name);
  }
});
