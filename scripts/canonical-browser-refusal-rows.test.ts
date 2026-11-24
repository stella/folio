import { expect, test } from "bun:test";

import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import {
  canonicalBrowserRefusalRows,
  matchCanonicalRefusalRow,
  validateHarnessRefusalRows,
} from "../test/canonical-refusal-rows";
import { BROWSER_PASTE_PAYLOADS } from "../tests/visual/browserInputTrace";
import type { BrowserInputAction } from "../tests/visual/browserInputTrace";

const tablePaste = {
  kind: "pasteTable",
  ...BROWSER_PASTE_PAYLOADS.pasteTable,
} satisfies BrowserInputAction;
const cells = { selection: "cells" } as const;
const imeOverCells = canonicalBrowserRefusalRows(
  { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
  cells,
);

test("a pasted table is strict at a caret and over selected cells", () => {
  expect(canonicalBrowserRefusalRows(tablePaste)).toEqual([]);
  expect(canonicalBrowserRefusalRows(tablePaste, cells)).toEqual([]);
});

test("canonical browser refusal matching takes a declared row and rejects unknown and mismatched refusals", () => {
  validateHarnessRefusalRows(imeOverCells);
  const declared = imeOverCells.at(0);
  if (declared === undefined)
    throw new TypeError("IME replacement over cells declares a refusal.");
  expect(
    matchCanonicalRefusalRow({ rows: imeOverCells, refusal: declared }),
  ).toEqual(declared);
  expect(
    matchCanonicalRefusalRow({
      rows: imeOverCells,
      refusal: { gap: CANONICAL_GAP.commands, message: "A different refusal." },
    }),
  ).toBeUndefined();
  expect(
    matchCanonicalRefusalRow({
      rows: imeOverCells,
      refusal: { gap: CANONICAL_GAP.commands, message: declared.message },
    }),
  ).toBeUndefined();
  expect(
    canonicalBrowserRefusalRows({ kind: "typing", text: "alpha" }),
  ).toEqual([]);
});

test("typing, pasting and deleting over selected cells stay strict; IME replacement declares its refusal", () => {
  // Typing and pasting clear every selected cell and land in the first.
  expect(
    canonicalBrowserRefusalRows(
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
      cells,
    ),
  ).toEqual([]);
  expect(
    canonicalBrowserRefusalRows({ kind: "typing", text: "alpha" }, cells),
  ).toEqual([]);
  for (const action of [
    { kind: "delete" },
    { kind: "backspace" },
    { kind: "cut" },
  ] as const)
    expect(canonicalBrowserRefusalRows(action, cells)).toEqual([]);
  expect(imeOverCells.length).toBeGreaterThan(0);
});

test("canonical browser refusal rows reject retired ledger gaps", () => {
  expect(() => validateHarnessRefusalRows(imeOverCells, {})).toThrow(
    "Retired canonical refusal row must become strict",
  );
});
