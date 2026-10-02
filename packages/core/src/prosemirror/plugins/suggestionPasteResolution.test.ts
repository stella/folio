import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { Fragment, Slice } from "prosemirror-model";

import { DOCUMENT_SHAPES } from "../../__tests__/documentShapes";
import {
  EXTRA_OPERATIONS,
  runConformanceCase,
  type ConformanceOperation,
} from "../../__tests__/editorCommandConformance";
import { TEXTBLOCK_SELECTION_PLACEMENTS } from "../../__tests__/editorHarness";

const copiedBlocks = EXTRA_OPERATIONS.find(({ id }) => id === "paste:copied-blocks");
const trackedShape = DOCUMENT_SHAPES.find(({ id }) => id === "tracked-changes");
if (!copiedBlocks || !trackedShape) panic("Missing paste resolution inputs");

describe("tracked paste resolution", () => {
  test.each(TEXTBLOCK_SELECTION_PLACEMENTS)(
    "copied revisions resolve like direct paste at %s",
    async (placement) => {
      const result = await runConformanceCase(trackedShape, copiedBlocks, placement);
      expect(result).not.toBeNull();
      expect(result?.violations).toEqual([]);
    },
  );
});

const REVISION_KINDS = ["insertion", "deletion", "moveTo", "moveFrom"] as const;
const nestedCases = REVISION_KINDS.flatMap((outer) =>
  REVISION_KINDS.flatMap((inner) =>
    TEXTBLOCK_SELECTION_PLACEMENTS.map((placement) => [outer, inner, placement] as const),
  ),
);
const plainShape = DOCUMENT_SHAPES.find(({ id }) => id === "bare-package");
if (!plainShape) panic("Missing plain paste resolution input");

// Cross the clipboard revision kinds and selection placements instead of
// allowing the known-gap catalog to mask a copied revision's resolution.
test.each(nestedCases)(
  "nested clipboard %s / %s resolves like direct paste at %s",
  async (outer, inner, placement) => {
    const operation = {
      id: "paste:nested-revisions",
      placements: TEXTBLOCK_SELECTION_PLACEMENTS,
      run: ({ view }) => {
        const markType =
          view.state.schema.marks[
            inner === "insertion" || inner === "moveTo" ? "insertion" : "deletion"
          ];
        if (!markType) panic("Missing clipboard revision mark");
        const mark = markType.create({
          revisionId: 200,
          author: "Inner",
          date: "2026-01-01T00:00:00Z",
          moveKind: inner === "moveTo" || inner === "moveFrom" ? inner : null,
          _docxRevisionAncestors: [
            {
              type: outer,
              revisionId: 100,
              author: "Outer",
              date: "2026-01-01T00:00:00Z",
              outerWrapperCount: 0,
            },
          ],
        });
        view.paste(new Slice(Fragment.from(view.state.schema.text("Clipboard", [mark])), 0, 0));
        return undefined;
      },
    } as const satisfies ConformanceOperation;
    const result = await runConformanceCase(plainShape, operation, placement);
    expect(result).not.toBeNull();
    expect(result?.violations).toEqual([]);
  },
);
