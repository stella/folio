import { describe, expect, test } from "bun:test";

import { emptyBaseline, findUntranslated, isSorted, sortKeys, syncMessages } from "./i18n-check";
import type { CheckBaseline, NestedMessages } from "./i18n-check";

describe("sortKeys", () => {
  test("sorts nested keys recursively", () => {
    const input: NestedMessages = {
      folio: { zoomGroup: "Zoom", clearDate: "Clear date" },
    };

    const result = sortKeys(input);

    const folio = result["folio"];
    expect(typeof folio === "object" && Object.keys(folio)).toEqual(["clearDate", "zoomGroup"]);
  });
});

describe("isSorted", () => {
  test("returns false when a trailing key is out of order (clearDate after zoomGroup)", () => {
    expect(isSorted({ folio: { zoomGroup: "Zoom", clearDate: "Clear date" } })).toBe(false);
  });

  test("returns true for a sorted nested catalog", () => {
    expect(isSorted({ folio: { clearDate: "Clear date", zoomGroup: "Zoom" } })).toBe(true);
  });
});

describe("syncMessages", () => {
  test("preserves translations while adding missing and dropping extra keys, sorted", () => {
    const source: NestedMessages = {
      folio: { insertTable: "Insert table", zoomGroup: "Zoom" },
    };
    const target: NestedMessages = {
      folio: { zoomGroup: "Priblíženie", obsolete: "Staré" },
    };

    expect(syncMessages(source, target)).toEqual({
      folio: { insertTable: "Insert table", zoomGroup: "Priblíženie" },
    });
  });

  test("keeps the new subtree when a key changes from a leaf to a namespace", () => {
    // en promoted `greeting` from a string to a namespace; the stale leaf must
    // drop without deleting the freshly-added `greeting.formal`.
    const source: NestedMessages = { greeting: { formal: "Good day" } };
    const target: NestedMessages = { greeting: "Hi" };

    expect(syncMessages(source, target)).toEqual({ greeting: { formal: "Good day" } });
  });
});

describe("findUntranslated", () => {
  const en: NestedMessages = {
    folio: {
      insertTable: "Insert table",
      zoomGroup: "Zoom",
      count: "{n}",
      ok: "OK",
      brand: "stella",
    },
  };

  test("flags a value that byte-equals the English source", () => {
    const target: NestedMessages = {
      folio: {
        insertTable: "Insert table",
        zoomGroup: "Priblíženie",
        count: "{n}",
        ok: "OK",
        brand: "stella",
      },
    };
    expect(findUntranslated(en, target, "sk", emptyBaseline())).toEqual(["folio.insertTable"]);
  });

  test("skips a real translation", () => {
    const target: NestedMessages = {
      folio: {
        insertTable: "Vložiť tabuľku",
        zoomGroup: "Priblíženie",
        count: "{n}",
        ok: "OK",
        brand: "stella",
      },
    };
    expect(findUntranslated(en, target, "sk", emptyBaseline())).toEqual([]);
  });

  test("respects ALLOWED_IDENTICAL and placeholder-only / brand values", () => {
    // "OK", "{n}", and lowercase "stella" are identical to en but must not be
    // flagged: OK/stella are in ALLOWED_IDENTICAL, {n} is placeholder-only.
    const target: NestedMessages = {
      folio: {
        insertTable: "Vložiť tabuľku",
        zoomGroup: "Priblíženie",
        count: "{n}",
        ok: "OK",
        brand: "stella",
      },
    };
    expect(findUntranslated(en, target, "sk", emptyBaseline())).toEqual([]);
  });

  test("respects the per-locale baseline grandfathering", () => {
    const target: NestedMessages = {
      folio: {
        insertTable: "Insert table",
        zoomGroup: "Zoom",
        count: "{n}",
        ok: "OK",
        brand: "stella",
      },
    };
    const baseline: CheckBaseline = {
      ...emptyBaseline(),
      identicalToSource: { "folio.insertTable": ["sk"] },
    };
    // Grandfathered for sk...
    expect(findUntranslated(en, target, "sk", baseline)).toEqual(["folio.zoomGroup"]);
    // ...but still flagged for a locale not in the baseline list.
    expect(findUntranslated(en, target, "de", baseline)).toEqual([
      "folio.insertTable",
      "folio.zoomGroup",
    ]);
  });
});
