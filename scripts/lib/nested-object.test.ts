import { afterEach, describe, expect, test } from "bun:test";

import { deleteNestedValue, getNestedValue, setNestedValue } from "./nested-object";
import type { NestedObject } from "./nested-object";

afterEach(() => {
  // SAFETY: test-only cleanup of intentional prototype pollution probes.
  Reflect.deleteProperty(Object.prototype as object, "inheritedLocaleBranch");
  Reflect.deleteProperty(Object.prototype as object, "polluted");
});

describe("setNestedValue", () => {
  test("writes own nested paths", () => {
    const target: NestedObject = {};
    setNestedValue(target, "folio.insertTable", "Insert table");
    expect(target).toEqual({ folio: { insertTable: "Insert table" } });
  });

  test("does not descend through inherited properties", () => {
    // SAFETY: temporary own-data probe on Object.prototype for this test only.
    Object.defineProperty(Object.prototype, "inheritedLocaleBranch", {
      value: {},
      writable: true,
      configurable: true,
      enumerable: false,
    });

    const target: NestedObject = {};
    setNestedValue(target, "inheritedLocaleBranch.leaf", "safe");

    expect(Object.hasOwn(target, "inheritedLocaleBranch")).toBe(true);
    expect(target.inheritedLocaleBranch).toEqual({ leaf: "safe" });
    expect((Object.prototype as { inheritedLocaleBranch?: unknown }).inheritedLocaleBranch).toEqual(
      {},
    );
  });

  test("rejects prototype-chain path segments", () => {
    const target: NestedObject = {};
    expect(() => setNestedValue(target, "constructor.prototype.polluted", "nope")).toThrow(
      /unsafe key path/,
    );
    expect((Object.prototype as { polluted?: unknown }).polluted).toBeUndefined();
  });
});

describe("deleteNestedValue", () => {
  test("removes a leaf and prunes empty parents", () => {
    const target: NestedObject = { folio: { insertTable: "Insert table" } };
    deleteNestedValue(target, "folio.insertTable");
    expect(target).toEqual({});
  });

  test("rejects prototype-chain path segments", () => {
    const target: NestedObject = {};
    expect(() => deleteNestedValue(target, "constructor.prototype.polluted")).toThrow(
      /unsafe key path/,
    );
    expect((Object.prototype as { polluted?: unknown }).polluted).toBeUndefined();
  });
});

describe("getNestedValue", () => {
  test("reads own nested paths", () => {
    const target: NestedObject = { folio: { insertTable: "Insert table" } };
    expect(getNestedValue(target, "folio.insertTable")).toBe("Insert table");
  });

  test("ignores inherited properties", () => {
    // SAFETY: temporary own-data probe on Object.prototype for this test only.
    Object.defineProperty(Object.prototype, "inheritedLocaleBranch", {
      value: { leaf: "inherited" },
      writable: true,
      configurable: true,
      enumerable: false,
    });

    const target: NestedObject = {};
    expect(getNestedValue(target, "inheritedLocaleBranch.leaf")).toBeUndefined();
  });
});
