import { expect, test } from "bun:test";
import { panic } from "better-result";

import {
  runConformanceCase,
  type ConformanceOperation,
} from "../../__tests__/editorCommandConformance";
import { textblocks } from "../../__tests__/editorHarness";
import { LIST_PASTE_RESOLUTION_CASES } from "../../__tests__/editorCommandConformance.listPaste";

// The matrix previously tolerated every paste resolution mismatch. Cross the
// direct-list shapes, clipboard kinds and selection placements independently
// of that catalog: accept equals direct paste, reject equals the original,
// and both outcomes keep their numbering after save and reopen.
test.each(
  LIST_PASTE_RESOLUTION_CASES.map(
    ({ shape, operation, placement }) =>
      [`${shape.id} / ${operation.id} / ${placement}`, { shape, operation, placement }] as const,
  ),
)(
  "tracked list paste resolves like direct paste: %s",
  async (_label, { shape, operation, placement }) => {
    const result = await runConformanceCase(shape, operation, placement);
    expect(result).not.toBeNull();
    expect(result?.violations).toEqual([]);
  },
);

// A single copied paragraph has no inserted internal break to carry the
// replaced paragraph's properties: the first-part change must restore them.
const singleParagraphPaste = {
  id: "paste:single-copied-paragraph",
  placements: [],
  run: ({ view, focus }) => {
    const { node, pos } =
      textblocks(view.state.doc).find(({ node: candidate }) => candidate.textContent === focus) ??
      panic("Missing copied paragraph");
    view.paste(view.state.doc.slice(pos + 1, pos + 1 + node.content.size));
    return undefined;
  },
} as const satisfies ConformanceOperation;

test.each(
  LIST_PASTE_RESOLUTION_CASES.filter(({ operation }) => operation.id === "paste:copied-blocks").map(
    ({ shape, placement }) => [`${shape.id} / ${placement}`, { shape, placement }] as const,
  ),
)("one copied paragraph restores its properties: %s", async (_label, { shape, placement }) => {
  const result = await runConformanceCase(shape, singleParagraphPaste, placement);
  expect(result).not.toBeNull();
  expect(result?.violations).toEqual([]);
});
