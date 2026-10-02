import { describe, expect, test } from "bun:test";

import { createModelVersionTracker } from "./modelVersion";

describe("model version tracker", () => {
  test("shared records are read once per traversal and mutations invalidate every alias", () => {
    const readVersion = createModelVersionTracker();
    let value = "before";
    let reads = 0;
    const shared = {
      get value() {
        reads += 1;
        return value;
      },
    };
    const model = { first: shared, nested: { second: shared }, items: [shared, shared] };
    const initial = readVersion(model);
    expect(reads).toBe(1);
    expect(readVersion(model)).toBe(initial);
    expect(reads).toBe(2);

    value = "after";
    const updated = readVersion(model);
    expect(updated).not.toBe(initial);
    expect(reads).toBe(3);
    expect(readVersion(model)).toBe(updated);
    expect(reads).toBe(4);

    model.nested.second = { value: "after" };
    expect(readVersion(model)).not.toBe(updated);
    expect(reads).toBe(5);
  });

  test("nested mutations produce a new version while unchanged reads reuse one", () => {
    const readVersion = createModelVersionTracker();
    const model = { nested: { value: "before" } };
    const initialVersion = readVersion(model);

    expect(readVersion(model)).toBe(initialVersion);
    model.nested.value = "after";
    const updatedVersion = readVersion(model);
    expect(updatedVersion).not.toBe(initialVersion);
    expect(readVersion(model)).toBe(updatedVersion);
  });

  test("array length changes and filling a hole produce new versions", () => {
    const readVersion = createModelVersionTracker();
    const items = ["first"];
    items.length = 3;
    items[2] = "third";
    const model = { items };
    const initialVersion = readVersion(model);

    items.length = 5;
    const extendedVersion = readVersion(model);
    expect(extendedVersion).not.toBe(initialVersion);

    items[1] = "second";
    const filledVersion = readVersion(model);
    expect(filledVersion).not.toBe(extendedVersion);
    expect(readVersion(model)).toBe(filledVersion);
  });

  test("deleting an object key produces a new version", () => {
    const readVersion = createModelVersionTracker();
    const model = { first: 1, second: 2 };
    const initialVersion = readVersion(model);

    Reflect.deleteProperty(model, "second");
    expect(readVersion(model)).not.toBe(initialVersion);
  });

  test("inherited enumerable getters are not read", () => {
    const readVersion = createModelVersionTracker();
    let inheritedReads = 0;
    const prototype = Object.defineProperty({}, "inherited", {
      enumerable: true,
      get: () => {
        inheritedReads += 1;
        return "inherited";
      },
    });
    const model = { own: "value" };
    Object.setPrototypeOf(model, prototype);

    const initial = readVersion(model);
    expect(readVersion(model)).toBe(initial);
    expect(inheritedReads).toBe(0);
  });

  test.each(["first", "second", "third"] as const)(
    "a warmed scan detects a mutation at the %s key",
    (key) => {
      const readVersion = createModelVersionTracker();
      const model = { first: "before", second: "before", third: "before" };
      const initial = readVersion(model);
      expect(readVersion(model)).toBe(initial);

      model[key] = "after";
      const updated = readVersion(model);
      expect(updated).not.toBe(initial);
      expect(readVersion(model)).toBe(updated);
    },
  );

  test("deleting and re-adding a key recognizes the changed own-key order", () => {
    const readVersion = createModelVersionTracker();
    const model = { first: 1, second: 2, third: 3 };
    const initial = readVersion(model);

    Reflect.deleteProperty(model, "second");
    Reflect.set(model, "second", 2);
    const reordered = readVersion(model);

    expect(reordered).not.toBe(initial);
    expect(readVersion(model)).toBe(reordered);
  });

  test("replacing a nested object by an equal object produces a new version", () => {
    const readVersion = createModelVersionTracker();
    const model = { nested: { value: "same" } };
    const initialVersion = readVersion(model);

    model.nested = { value: "same" };
    const replacementVersion = readVersion(model);
    expect(replacementVersion).not.toBe(initialVersion);
    expect(readVersion(model)).toBe(replacementVersion);
  });

  test("adding undefined fields and removing trailing keys invalidate a warmed version", () => {
    const readVersion = createModelVersionTracker();
    const model = { first: "value" };
    const initial = readVersion(model);
    Reflect.set(model, "last", undefined);
    const expanded = readVersion(model);
    expect(expanded).not.toBe(initial);
    expect(readVersion(model)).toBe(expanded);
    Reflect.deleteProperty(model, "last");
    const reduced = readVersion(model);
    expect(reduced).not.toBe(expanded);
    expect(readVersion(model)).toBe(reduced);
  });
});
