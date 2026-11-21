import { Glob } from "bun";
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { parseSourceModule, walkAst } from "../../../../../scripts/lib/oxc-ast";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { isOwnRecord } from "../revisionOwner";
import { withoutPendingSuggestions } from "../suggestions";
import { OP_STORIES } from "../types";

const author = "Drafter";
const inserted = (paraId: string, text: string, id: number): Paragraph => ({
  type: "paragraph",
  paraId,
  pPrMark: { kind: "ins", info: { id, author, date: "2026-01-01T00:00:00Z" } },
  content: [
    {
      type: "insertion",
      info: { id: id + 1, author, date: "2026-01-01T00:00:00Z" },
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});
const document: Document = {
  package: {
    document: {
      content: [
        inserted("00000001", "An inserted clause.", 1),
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            { type: "run", content: [{ type: "text", text: "Next." }] },
          ],
        },
      ],
    },
  },
};
const at = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

const INTENTS: Record<string, EditorIntent> = {
  replaceText: {
    type: "replaceText",
    from: at("00000001", 3),
    to: at("00000001", 11),
    text: "added",
  },
  deleteText: {
    type: "replaceText",
    from: at("00000001", 3),
    to: at("00000001", 11),
    text: "",
  },
  deleteParagraph: {
    type: "deleteParagraphs",
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
  },
};

const suggest = (intent: EditorIntent, suggestionId?: string) => {
  const compiled = compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: {
        id: 10,
        author,
        date: "2026-01-02T00:00:00Z",
        ...(suggestionId === undefined ? {} : { suggestionId }),
      },
      newIds: { revision: [11, 12, 13, 14, 15, 16] },
    },
  }).unwrap();
  return applyDocumentOps(document, compiled.ops).unwrap().document;
};

// A suggestion takes back only what it recorded itself: over its author's
// ordinary pending insertion it records its deletion, which rejecting the
// suggestion gives back, as a save without pending suggestions does.
for (const [name, intent] of Object.entries(INTENTS))
  test(`a suggested ${name} over its author's ordinary insertion is rejectable`, () => {
    expect(withoutPendingSuggestions(suggest(intent, "s1")).unwrap()).toEqual(
      document,
    );
  });

test("a record is its owner's only with the same author and suggestion", () => {
  const info = { author, suggestionId: "s1" };
  expect(isOwnRecord(info, { author, suggestionId: "s1" })).toBe(true);
  expect(isOwnRecord(info, { author, suggestionId: "s2" })).toBe(false);
  expect(isOwnRecord(info, { author })).toBe(false);
  expect(isOwnRecord({ author }, { author, suggestionId: "s1" })).toBe(false);
  expect(isOwnRecord({ author }, { author: "Reviewer" })).toBe(false);
  expect(isOwnRecord({ author }, { author })).toBe(true);
});

type AstNode = { type: string } & Record<string, unknown>;
const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value;
/** `x.author` or `Reflect.get(x, "author")`. */
const readsAuthor = (node: unknown): boolean => {
  if (!isNode(node)) return false;
  if (
    node.type === "MemberExpression" ||
    node.type === "StaticMemberExpression"
  )
    return isNode(node["property"]) && node["property"]["name"] === "author";
  if (node.type !== "CallExpression") return false;
  const args = node["arguments"];
  return (
    Array.isArray(args) &&
    isNode(args[1]) &&
    args[1]["value"] === "author" &&
    isNode(node["callee"]) &&
    isNode(node["callee"]["property"]) &&
    node["callee"]["property"]["name"] === "get"
  );
};
const isAbsence = (node: unknown): boolean =>
  isNode(node) &&
  ((node.type === "Literal" && node["value"] === null) ||
    (node.type === "Identifier" && node["name"] === "undefined"));
const authorComparisons = (source: string): number => {
  const { program } = parseSourceModule("source.ts", source);
  let count = 0;
  walkAst(program, (node) => {
    if (
      node.type === "BinaryExpression" &&
      (node.operator === "===" || node.operator === "!==") &&
      (readsAuthor(node.left) || readsAuthor(node.right)) &&
      !isAbsence(node.left) &&
      !isAbsence(node.right)
    )
      count++;
  });
  return count;
};

test("author comparisons detect member reads and reflected reads", () => {
  expect(authorComparisons("a.info.author === b.author")).toBe(1);
  expect(authorComparisons('Reflect.get(info, "author") !== author')).toBe(1);
  expect(authorComparisons("op.patch.author === null")).toBe(0);
  expect(authorComparisons('typeof a.author === "string"')).toBe(0);
});

// Whether an edit takes back a record as its own is one decision: same author
// and same suggestion. Each operation asks `isOwnRecord`, never the author alone.
test("every own-record decision of the operations goes through isOwnRecord", () => {
  const root = path.join(import.meta.dir, "..");
  const files = [
    ...new Glob("**/*.ts").scanSync({ cwd: root, onlyFiles: true }),
  ].filter(
    (file) =>
      file !== "revisionOwner.ts" &&
      !file.includes("__tests__/") &&
      !file.endsWith(".test.ts") &&
      !file.endsWith(".typecheck.ts"),
  );
  expect(files.length).toBeGreaterThan(0);
  expect(
    files.filter(
      (file) =>
        authorComparisons(readFileSync(path.join(root, file), "utf8")) > 0,
    ),
  ).toEqual([]);
});
