/**
 * Every editor operation — each registry command, key binding, paste and host
 * flow the conformance matrix drives — must build what it inserts from the
 * document's own schema.
 *
 * Header, footer and note editors run an extension runtime of their own, built
 * on a schema instance other than the one documents are parsed into. An
 * extension that creates nodes or marks from types it captured when it was set
 * up, rather than from `state.schema`, then builds content that document cannot
 * hold: the step fails to fit and the edit silently does nothing or half of
 * what it should, or ProseMirror throws. So each operation runs twice on every
 * shape, once on the runtime built with the document's schema and once on a
 * runtime built with another, and the two must end alike.
 */

import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import { DOCUMENT_SHAPES } from "./documentShapes";
import { CONFORMANCE_OPERATIONS, type ConformanceOperation } from "./editorCommandConformance";
import {
  createHarnessState,
  EDITOR_MODES,
  type EditorMode,
  type HarnessRuntime,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  SELECTION_PLACEMENTS,
  type SelectionPlacement,
} from "./editorHarness";
import type { Document } from "../types/document";

/** Attributes minted per run (revision ids and dates, fresh paragraph ids). */
const VOLATILE_KEYS = new Set([
  "date",
  "utcDate",
  "revisionId",
  "id",
  "paraId",
  "textId",
  "_docxParagraphSourceToken",
]);

const stable = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    return value.map(stable);
  }
  if (typeof value === "object" && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([key]) => !VOLATILE_KEYS.has(key))
        .map(([key, entry]) => [key, stable(entry)]),
    );
  }
  // A table of contents names its heading bookmarks at random.
  return typeof value === "string" ? value.replaceAll(/_Toc\d+/gu, "_Toc") : value;
};

const firstDifference = (left: string, right: string): string => {
  let index = 0;
  while (index < left.length && left[index] === right[index]) {
    index += 1;
  }
  return `…${left.slice(Math.max(0, index - 80), index + 60)}… ≠ …${right.slice(Math.max(0, index - 80), index + 60)}…`;
};

type Outcome = { status: string; doc: string };

const outcome = (
  base: Document,
  focus: string,
  operation: ConformanceOperation,
  placement: SelectionPlacement,
  mode: EditorMode,
  runtime: HarnessRuntime,
): Outcome | null => {
  const before: EditorState | null = placeSelection(
    createHarnessState(base, mode, [], runtime),
    focus,
    placement,
  );
  if (!before) {
    return null;
  }
  const view = new HeadlessEditorView(before);
  const caseBase: Document = {
    ...base,
    package: {
      ...base.package,
      document: {
        ...base.package.document,
        ...(base.package.document.comments
          ? { comments: [...base.package.document.comments] }
          : {}),
      },
    },
  };
  let status: string;
  try {
    const verdict = operation.run({ view, base: caseBase, focus });
    if (!view.state.doc.eq(before.doc)) {
      status = "changed";
    } else {
      status = verdict === false ? "refused" : "unchanged";
    }
  } catch (error) {
    status = `threw ${error instanceof Error ? error.message : String(error)}`;
  }
  return { status, doc: JSON.stringify(stable(view.state.doc.toJSON())) };
};

describe("operations build from the document's schema", () => {
  test.each(DOCUMENT_SHAPES.map((shape) => [shape.id, shape] as const))(
    "%s",
    async (_id, shape) => {
      const base = await parseShapeDocument(await shape.build());
      const differences: string[] = [];
      for (const operation of CONFORMANCE_OPERATIONS) {
        for (const placement of SELECTION_PLACEMENTS) {
          for (const mode of EDITOR_MODES) {
            const home = outcome(base, shape.focus, operation, placement, mode, "document");
            const foreign = outcome(base, shape.focus, operation, placement, mode, "harness");
            if (!home || !foreign) {
              continue;
            }
            if (home.status !== foreign.status || home.doc !== foreign.doc) {
              differences.push(
                `${operation.id} @ ${placement} [${mode}]: ${home.status} with the document's schema, ${foreign.status} with another${home.status === foreign.status ? ` (${firstDifference(home.doc, foreign.doc)})` : ""}`,
              );
            }
          }
        }
      }
      expect(differences).toEqual([]);
    },
    120_000,
  );
});
