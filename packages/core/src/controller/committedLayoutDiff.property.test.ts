/**
 * Property: an incremental pass may keep a committed measure only for a block
 * the committed and next documents share. Every block the derived range leaves
 * untouched must be the same block, at the same index, in both documents,
 * whatever edits (at any number of places) separate them.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));

import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { findDirtyBlockIndexes } from "../paged-layout/incrementalMeasure";
import { schema } from "../prosemirror/schema";
import { diffAgainstCommittedLayout, type LayoutFlowSource } from "./committedLayoutDiff";

const NO_PREVIEW = { entries: [], hidden: [], mode: "plain" } as const;

const source = (doc: PMNode): LayoutFlowSource => ({ doc, preview: NO_PREVIEW });

const paragraphText = fc
  .array(fc.constantFrom("a", "b", "ab", "the", " "), { maxLength: 12 })
  .map((parts) => parts.join(""));

const docFrom = (paragraphs: readonly string[]): PMNode =>
  schema.node(
    "doc",
    null,
    paragraphs.map((text) =>
      schema.node("paragraph", null, text.length > 0 ? [schema.text(text)] : []),
    ),
  );

type Edit = { type: "insert" | "delete"; at: number; text: string; length: number };

const edit: fc.Arbitrary<Edit> = fc.record({
  type: fc.constantFrom("insert" as const, "delete" as const),
  at: fc.nat(400),
  text: paragraphText.filter((text) => text.length > 0),
  length: fc.integer({ min: 1, max: 10 }),
});

/** Apply edits inside paragraphs only, so the block count stays the same. */
const applyEdits = (doc: PMNode, edits: readonly Edit[]): PMNode => {
  let state = EditorState.create({ doc });
  for (const { type, at, text, length } of edits) {
    const textPositions: number[] = [];
    state.doc.forEach((node, offset) => {
      for (let pos = offset + 1; pos <= offset + node.nodeSize - 1; pos += 1) {
        textPositions.push(pos);
      }
    });
    const from = textPositions[at % textPositions.length];
    if (from === undefined) {
      continue;
    }
    if (type === "insert") {
      state = state.apply(state.tr.insertText(text, from));
      continue;
    }
    const $from = state.doc.resolve(from);
    const to = Math.min($from.end(), from + length);
    if (to > from) {
      state = state.apply(state.tr.delete(from, to));
    }
  }
  return state.doc;
};

describe("diffAgainstCommittedLayout (properties)", () => {
  test("every block outside the derived range is unchanged at its index", () => {
    fc.assert(
      fc.property(
        fc.array(paragraphText, { minLength: 1, maxLength: 8 }),
        fc.array(edit, { minLength: 1, maxLength: 6 }),
        (paragraphs, edits) => {
          const committed = docFrom(paragraphs);
          const next = applyEdits(committed, edits);
          const diff = diffAgainstCommittedLayout(source(committed), source(next));
          if (diff.type === "full") {
            // Only an unchanged document may fall back to a full measure here.
            expect(next.eq(committed)).toBe(true);
            return;
          }
          const dirty = new Set(findDirtyBlockIndexes(toFlowBlocks(next), diff.range));
          for (let index = 0; index < next.childCount; index += 1) {
            if (!dirty.has(index)) {
              expect(next.child(index).eq(committed.child(index))).toBe(true);
            }
          }
        },
      ),
      propertyConfig({ numRuns: 500 }),
    );
  });

  test("a change to document attributes measures everything", () => {
    const committed = docFrom(["a", "b"]);
    const next = committed.type.create(
      { ...committed.attrs, _finalSectionStart: "continuous" },
      committed.content,
    );
    expect(diffAgainstCommittedLayout(source(committed), source(next))).toEqual({ type: "full" });
  });
});

describe("diffAgainstCommittedLayout (template preview)", () => {
  // "Intro {{ terms }} end." then "Tail."; the marker spans 7..18.
  const committedDoc = docFrom(["Intro {{ terms }} end.", "Tail."]);
  const marker = { from: 7, to: 18, expr: "terms" };
  const typedBefore = applyEdits(committedDoc, [{ type: "insert", at: 0, text: "ab", length: 1 }]);
  const preview = (
    entries: readonly { from: number; to: number; expr: string; value: string }[],
  ): LayoutFlowSource["preview"] => ({ entries, hidden: [], mode: "plain" });

  test("a marker only shifted by an edit before it is not dirty", () => {
    const diff = diffAgainstCommittedLayout(
      { doc: committedDoc, preview: preview([{ ...marker, value: "x" }]) },
      { doc: typedBefore, preview: preview([{ from: 9, to: 20, expr: "terms", value: "x" }]) },
    );
    expect(diff).toEqual({ type: "range", range: { from: 1, to: 3 } });
  });

  test("a marker that lost its value behind an edit is dirty where it now is", () => {
    const diff = diffAgainstCommittedLayout(
      { doc: committedDoc, preview: preview([{ ...marker, value: "x" }]) },
      { doc: typedBefore, preview: preview([]) },
    );
    expect(diff).toEqual({ type: "range", range: { from: 1, to: 20 } });
  });

  test("a preview mode switch measures everything", () => {
    const diff = diffAgainstCommittedLayout(
      { doc: committedDoc, preview: preview([]) },
      { doc: committedDoc, preview: { entries: [], hidden: [], mode: "highlighted" } },
    );
    expect(diff).toEqual({ type: "full" });
  });
});
