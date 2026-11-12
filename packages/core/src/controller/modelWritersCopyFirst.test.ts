/**
 * A writer that starts from a host's document model copies it first. Each
 * test here deep-freezes a model, runs one writer a host calls on it, and
 * checks that the writer neither throws on the frozen input nor returns an
 * object it shares with that input: the result is the caller's to write.
 */

import { describe, expect, test } from "bun:test";

import { deepFreeze } from "../__tests__/deepFreeze";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import {
  cloneCommentsWithParagraphPropertySources,
  cloneDocumentWithParagraphPropertySources,
} from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { FindReplaceManager } from "../managers/FindReplaceManager";
import { TableSelectionManager } from "../managers/TableSelectionManager";
import type { ReadonlyDocument } from "../types/documentModel";
import { findInDocument } from "../utils/findReplace";
import type { TableAction } from "../utils/tableOperations";

const shapeDocument = async (id: string): Promise<ReadonlyDocument> => {
  const shape = DOCUMENT_SHAPES.find((candidate) => candidate.id === id);
  if (!shape) throw new Error(`No document shape ${id}`);
  const document = await parseDocx(
    (
      await prepareCanonicalDocxInput((await shape.build()).slice().buffer)
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
  return deepFreeze(cloneDocumentWithParagraphPropertySources(document));
};

/**
 * Every object of document data reachable from `value`. Symbol-keyed handles
 * are left out: they carry a paragraph's parse provenance, which the copy
 * owner hands on to the copy on purpose, and hold no document content.
 */
const reachable = (value: unknown, seen = new Set<object>()): Set<object> => {
  if (typeof value !== "object" || value === null || seen.has(value))
    return seen;
  seen.add(value);
  if (value instanceof Map) {
    for (const [key, entry] of value) {
      reachable(key, seen);
      reachable(entry, seen);
    }
  } else if (value instanceof Set) {
    for (const entry of value) reachable(entry, seen);
  } else if (!(ArrayBuffer.isView(value) || value instanceof ArrayBuffer)) {
    for (const entry of Object.values(value)) reachable(entry, seen);
  }
  return seen;
};

/** The objects `result` shares with `input`. */
const sharedObjects = (input: unknown, result: unknown): object[] => {
  const inputs = reachable(input);
  return [...reachable(result)].filter((object) => inputs.has(object));
};

const STRUCTURAL_TABLE_ACTIONS = [
  "addRowAbove",
  "addRowBelow",
  "addColumnLeft",
  "addColumnRight",
  "deleteRow",
  "deleteColumn",
  "mergeCells",
  "splitCell",
  "deleteTable",
] as const satisfies readonly TableAction[];

describe("writers copy the model first", () => {
  test("the document copy owner", async () => {
    const model = await shapeDocument("comments");
    const copy = cloneDocumentWithParagraphPropertySources(model);
    expect(copy).toEqual(cloneDocumentWithParagraphPropertySources(model));
    expect(sharedObjects(model, copy)).toEqual([]);
  });

  test("the comment copy owner", async () => {
    const model = await shapeDocument("comments");
    const comments = model.package.document.comments;
    expect(comments?.length).toBeGreaterThan(0);
    const copy = cloneCommentsWithParagraphPropertySources(comments);
    expect(copy).toEqual(cloneCommentsWithParagraphPropertySources(comments));
    expect(copy.length).toBe(comments?.length ?? 0);
    expect(sharedObjects(model, copy)).toEqual([]);
  });

  test("find and replace, one match and all of them", async () => {
    const model = await shapeDocument("tables");
    const options = { matchCase: false, matchWholeWord: false };
    const plain = cloneDocumentWithParagraphPropertySources(model);
    const matches = findInDocument(plain, "e", options);
    expect(matches.length).toBeGreaterThan(1);

    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    const one = manager.replaceCurrent(model, "E");
    expect(one).not.toBeNull();
    expect(sharedObjects(model, one)).toEqual([]);

    const all = manager.replaceAll(model, matches, "E");
    expect(all?.replacedCount).toBeGreaterThan(0);
    expect(sharedObjects(model, all?.document)).toEqual([]);
  });

  test.each(STRUCTURAL_TABLE_ACTIONS)("the table action %s", async (action) => {
    const model = await shapeDocument("tables");
    const manager = new TableSelectionManager();
    expect(
      manager.selectCell(model, { tableIndex: 0, rowIndex: 0, columnIndex: 0 }),
    ).not.toBeNull();
    const result = manager.handleAction(model, action);
    if (result.type === "noop") return;
    expect(sharedObjects(model, result.document)).toEqual([]);
  });
});
