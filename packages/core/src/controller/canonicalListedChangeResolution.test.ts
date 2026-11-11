/**
 * A host resolves the changes the review list shows, by the ids it lists. A
 * tracked deletion beside a comment can be one listed change over several
 * revision ids, so resolving that change in the canonical session reaches all
 * of them: rejecting every listed change restores the source, and accepting or
 * rejecting them all leaves nothing pending, after save and reopen.
 */
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { AllSelection, TextSelection } from "prosemirror-state";

import {
  createCanonicalHarnessCase,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { modelMarkdown, parseShapeDocument } from "../__tests__/editorHarness";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { Document } from "../types/document";

type CutRange = { type: "all" } | { type: "text"; from: number; to: number };

const isMarker = (node: PMNode) =>
  node.type.name === "commentReference" ||
  node.type.name === "rangeAnchor" ||
  node.marks.some(({ type }) => type.name === "comment");

/**
 * Every range whose ends sit at inline boundaries, one character inside a
 * text, or a paragraph's edges, and which touches a comment marker.
 */
const rangesBesideComments = (doc: PMNode): CutRange[] => {
  const ends = new Set<number>();
  const markers = new Set<number>();
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      ends.add(pos + 1);
      ends.add(pos + node.nodeSize - 1);
      return true;
    }
    if (!node.isInline) return true;
    const end = pos + node.nodeSize;
    ends.add(pos);
    ends.add(end);
    if (node.isText && node.nodeSize > 2) {
      ends.add(pos + 1);
      ends.add(end - 1);
    }
    if (isMarker(node)) {
      markers.add(pos);
      markers.add(end);
    }
    return false;
  });
  const sorted = [...ends].sort((left, right) => left - right);
  const ranges: CutRange[] = [{ type: "all" }];
  for (const [index, from] of sorted.entries())
    for (const to of sorted.slice(index + 1))
      if ([...markers].some((marker) => from <= marker && marker <= to))
        ranges.push({ type: "text", from, to });
  return ranges;
};

const ready = (document: Document, mode: "editing" | "suggesting") => {
  const created = createCanonicalHarnessCase(document, mode);
  return created.type === "ready" ? created.driver : null;
};

/** The saved package after a suggested cut over `range`, or null when refused. */
const cutAndSave = async (bytes: Uint8Array, range: CutRange) => {
  const driver = ready(await parseShapeDocument(bytes), "suggesting");
  if (driver === null) throw new TypeError("Canonical cut unavailable.");
  try {
    const doc = driver.state.doc;
    driver.dispatch(
      driver.state.tr.setSelection(
        range.type === "all"
          ? new AllSelection(doc)
          : TextSelection.create(doc, range.from, range.to),
      ),
    );
    driver.cut();
    // A refused cut leaves the document as it was; nothing to resolve.
    if (driver.refusals.length > 0) return null;
    return (await saveCanonicalHarnessDocument(driver.snapshot())).bytes;
  } finally {
    driver.dispose();
  }
};

/** The saved package after a host resolves `ids`, or null when refused. */
const resolveAndSave = async (
  bytes: Uint8Array,
  ids: readonly number[],
  decision: "accept" | "reject",
) => {
  const driver = ready(await parseShapeDocument(bytes), "editing");
  if (driver === null) throw new TypeError("Canonical review unavailable.");
  try {
    if (!driver.history.resolveCanonicalRevisions(ids, decision)) return null;
    return (await saveCanonicalHarnessDocument(driver.snapshot())).bytes;
  } finally {
    driver.dispose();
  }
};

const listedChanges = async (bytes: Uint8Array) =>
  (await FolioDocxReviewer.fromBuffer(bytes.slice().buffer)).getChanges();

const describeRange = (range: CutRange) =>
  range.type === "all" ? "everything" : `${range.from}..${range.to}`;

const COMMENT_SHAPES = DOCUMENT_SHAPES.filter(({ features }) =>
  features.includes("comment"),
);

describe("resolving the listed changes of a suggested cut beside a comment", () => {
  test("covers at least one shape", () => {
    expect(COMMENT_SHAPES.length).toBeGreaterThan(0);
  });

  for (const shape of COMMENT_SHAPES) {
    test(
      shape.id,
      async () => {
        const bytes = await shape.build();
        const source = await parseShapeDocument(bytes);
        const probe = ready(source, "suggesting");
        if (probe === null) return;
        const ranges = rangesBesideComments(probe.state.doc);
        probe.dispose();
        const baseline = modelMarkdown(
          await parseShapeDocument(
            (await saveCanonicalHarnessDocument(source)).bytes,
          ),
        );
        const failures: string[] = [];
        let resolvedCases = 0;
        for (const range of ranges) {
          const saved = await cutAndSave(bytes, range);
          if (saved === null) {
            if (range.type === "all") failures.push("everything: cut refused");
            continue;
          }
          const listed = (await listedChanges(saved)).map(({ id }) => id);
          if (listed.length === 0) {
            if (range.type === "all")
              failures.push("everything: the cut listed no change");
            continue;
          }
          resolvedCases += 1;
          for (const decision of ["accept", "reject"] as const) {
            const label = `${describeRange(range)} ${decision}`;
            const resolved = await resolveAndSave(saved, listed, decision);
            if (resolved === null) {
              failures.push(`${label}: refused ${JSON.stringify(listed)}`);
              continue;
            }
            const left = await listedChanges(resolved);
            if (left.length > 0)
              failures.push(
                `${label}: ${JSON.stringify(listed)} left ${JSON.stringify(left.map(({ id, type, text }) => ({ id, type, text })))}`,
              );
            if (
              decision === "reject" &&
              modelMarkdown(await parseShapeDocument(resolved)) !== baseline
            )
              failures.push(`${label}: the source did not come back`);
          }
        }
        expect(failures).toEqual([]);
        expect(resolvedCases).toBeGreaterThan(0);
      },
      240_000,
    );
  }
});
