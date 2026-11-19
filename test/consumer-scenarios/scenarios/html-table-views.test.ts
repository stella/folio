import assert from "node:assert/strict";
import { test } from "node:test";

import { htmlTableViews } from "../support/html-table-views.ts";

// A string-only fixture missed XML parser coercion of numeric and boolean cells.
for (const text of [
  "0",
  "0007",
  "-12",
  "1.50",
  "1e3",
  "true",
  "false",
  "NaN",
  "Infinity",
]) {
  test(`HTML table reader preserves literal cell text ${text}`, () => {
    const expected = [{ text, kind: "paragraph" }];
    for (const content of [
      text,
      `<strong>${text}</strong>`,
      `<u>${text}</u>`,
    ]) {
      assert.deepEqual(
        htmlTableViews(`<table><tr><td>${content}</td></tr></table>`),
        expected,
      );
    }
  });
}

test("a table nested in a cell reads as its cells after the cell's text", () => {
  assert.deepEqual(
    htmlTableViews(
      "<table><tr><td>Outer<table><tr><td>Inner A</td><td>Inner B</td></tr></table></td><td>Next</td></tr></table>",
    ),
    [
      { text: "Outer", kind: "paragraph" },
      { text: "Inner A", kind: "paragraph" },
      { text: "Inner B", kind: "paragraph" },
      { text: "Next", kind: "paragraph" },
    ],
  );
});
