import { expect, test } from "bun:test";

import {
  canonicalBrowserRefusalRows,
  matchCanonicalRefusalRow,
  validateHarnessRefusalRows,
} from "../test/canonical-refusal-rows";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import { BROWSER_PASTE_PAYLOADS } from "../tests/visual/browserInputTrace";
import type { BrowserInputAction } from "../tests/visual/browserInputTrace";

const tablePaste = {
  kind: "pasteTable",
  ...BROWSER_PASTE_PAYLOADS.pasteTable,
} satisfies BrowserInputAction;

test("canonical browser table paste declares exact refusal alternatives", () => {
  const rows = canonicalBrowserRefusalRows(tablePaste);
  expect(rows.map(({ id }) => id)).toEqual(["clipboard-table", "clipboard-nested-blocks"]);
  validateHarnessRefusalRows(rows);

  expect(
    matchCanonicalRefusalRow({
      rows,
      refusal: {
        gap: CANONICAL_GAP.dispatch,
        message: "Clipboard tables and embedded blocks require canonical table editing.",
      },
    }),
  ).toEqual(rows[0]);
  expect(
    matchCanonicalRefusalRow({
      rows,
      refusal: {
        gap: CANONICAL_GAP.dispatch,
        message: "Clipboard tables and nested block containers require canonical table editing.",
      },
    }),
  ).toEqual(rows[1]);
});

test("canonical browser refusal matching rejects unknown and mismatched refusals", () => {
  const rows = canonicalBrowserRefusalRows(tablePaste);
  expect(
    matchCanonicalRefusalRow({
      rows,
      refusal: { gap: CANONICAL_GAP.dispatch, message: "A different refusal." },
    }),
  ).toBeUndefined();
  expect(
    matchCanonicalRefusalRow({
      rows,
      refusal: {
        gap: CANONICAL_GAP.tableActivation,
        message: "Clipboard tables and embedded blocks require canonical table editing.",
      },
    }),
  ).toBeUndefined();
  expect(canonicalBrowserRefusalRows({ kind: "typing", text: "alpha" })).toEqual([]);
});

test("canonical browser refusal rows reject retired ledger gaps", () => {
  const rows = canonicalBrowserRefusalRows(tablePaste);
  expect(() => validateHarnessRefusalRows(rows, {})).toThrow(
    "Retired canonical refusal row must become strict",
  );
});
