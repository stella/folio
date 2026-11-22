import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  compileEditorIntent,
  type EditorIntent,
  type EditorIntentMode,
} from "../editorIntent";
import { paragraphLength } from "../offsets";
import { isBlockBoundary } from "../rangeBoundaries";
import { paragraphPropertiesOf } from "../review";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

// A paragraph whose mark a pending deletion takes is one paragraph with the
// next: deleting the break after the two gives the words those paragraphs
// hold together the properties of the last one, whether that one holds words
// itself, and whether the deletion is direct or tracked and then accepted.

const earlier = (id: number) => ({
  id,
  author: "Earlier",
  date: "2026-01-02T03:04:05Z",
});
const REVISION = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };

/** A paragraph's words, then `D` for a mark pending deletion, then its style. */
type Spec = `${string}${"" | "D"}:${string}`;

const documentOf = (specs: readonly Spec[]): Document => ({
  package: {
    document: {
      content: specs.map((spec, index): Paragraph => {
        const [state = "", styleId = ""] = spec.split(":");
        const pending = state.endsWith("D");
        const text = pending ? state.slice(0, -1) : state;
        const paragraph: Paragraph = {
          type: "paragraph",
          paraId: `0000000${index + 1}`,
          formatting: { styleId },
          content:
            text === ""
              ? []
              : [{ type: "run", content: [{ type: "text", text }] }],
        };
        if (pending)
          paragraph.pPrMark = {
            kind: "del",
            info: earlier(index + 1),
          };
        return paragraph;
      }),
    },
  },
});

const paragraphsOf = (document: Document): Paragraph[] =>
  document.package.document.content
    .filter((block) => !isBlockBoundary(block))
    .map((block) =>
      block.type === "paragraph" ? block : panic("A test body is paragraphs."),
    );

/** Every revision id a document holds. */
const revisionIds = (document: Document): number[] => {
  const ids: number[] = [];
  for (const paragraph of paragraphsOf(document)) {
    if (paragraph.pPrMark !== undefined) ids.push(paragraph.pPrMark.info.id);
    for (const { info } of paragraph.propertyChanges ?? []) ids.push(info.id);
    for (const item of paragraph.content)
      if (item.type === "deletion") ids.push(item.info.id);
  }
  return ids;
};

const resolve = (
  document: Document,
  ids: readonly number[],
  decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
): Document =>
  ids.length === 0
    ? document
    : applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [...ids],
        decision,
      }).unwrap().document;

const view = (document: Document): string[] =>
  paragraphsOf(document).map((paragraph) => {
    const text = paragraph.content
      .flatMap((item) => (item.type === "run" ? item.content : []))
      .flatMap((leaf) => (leaf.type === "text" ? [leaf.text] : []))
      .join("");
    return `${text}:${paragraphPropertiesOf(paragraph.formatting)?.styleId}`;
  });

const MODES = {
  direct: { type: "editing" },
  tracked: {
    type: "suggesting",
    revision: REVISION,
    newIds: { revision: Array.from({ length: 64 }, (_, index) => 101 + index) },
  },
} satisfies Record<string, EditorIntentMode>;

/** The view once `intent` is applied in `mode` and every change accepted. */
const accepted = (
  document: Document,
  intent: EditorIntent,
  mode: EditorIntentMode,
): string[] => {
  const { ops } = compileEditorIntent(document, { intent, mode }).unwrap();
  const applied = applyDocumentOps(document, ops).unwrap();
  expect(
    applyDocumentOps(applied.document, applied.inverse).unwrap().document,
  ).toEqual(document);
  if (mode.type === "suggesting")
    expect(
      resolve(
        applied.document,
        revisionIds(applied.document).filter(
          (id) => !revisionIds(document).includes(id),
        ),
        REVISION_DECISIONS.REJECT,
      ),
    ).toEqual(document);
  return view(
    resolve(
      applied.document,
      revisionIds(applied.document),
      REVISION_DECISIONS.ACCEPT,
    ),
  );
};

/** The break after `specs[from]`, deleted as each intent deletes it. */
const intentsAt = (document: Document, from: number) => {
  const story = OP_STORIES.MAIN;
  const [first, next] = paragraphsOf(document).slice(from, from + 2);
  if (first === undefined || next === undefined)
    return panic("A break stands between two paragraphs.");
  const blockId = first.paraId ?? "";
  const nextBlockId = next.paraId ?? "";
  return {
    range: {
      type: "replaceText",
      from: { story, blockId, offset: paragraphLength(first) },
      to: { story, blockId: nextBlockId, offset: paragraphLength(next) },
      text: "",
    },
    join: { type: "joinParagraphs", story, blockId, nextBlockId },
    deleteNext: { type: "deleteParagraphs", story, blockIds: [nextBlockId] },
  } satisfies Record<string, EditorIntent>;
};

type Case = {
  specs: readonly Spec[];
  from: number;
  intent: keyof ReturnType<typeof intentsAt>;
  expected: readonly string[];
};

const cases: Record<string, Case> = {
  "a range over the break of an empty paragraph words run into": {
    specs: ["aD:P", ":Q", ":R"],
    from: 1,
    intent: "range",
    expected: ["a:Q"],
  },
  "a join of an empty paragraph words run into": {
    specs: ["aD:P", ":Q", "r:R"],
    from: 1,
    intent: "join",
    expected: ["ar:Q"],
  },
  "a last paragraph after an empty one words run into": {
    specs: ["aD:P", ":Q", "r:R"],
    from: 1,
    intent: "deleteNext",
    expected: ["a:Q"],
  },
  "words run in through several pending marks": {
    specs: ["aD:P", "D:O", ":Q", "r:R"],
    from: 2,
    intent: "range",
    expected: ["a:Q"],
  },
  "no words run into an empty paragraph": {
    specs: ["D:P", ":Q", ":R"],
    from: 1,
    intent: "range",
    expected: [":R"],
  },
  "a range before a paragraph that stays": {
    specs: ["aD:P", ":Q", "r:R", "z:Z"],
    from: 1,
    intent: "range",
    expected: ["a:Q", "z:Z"],
  },
  "a range over a break already pending deletion": {
    specs: ["aD:P", ":Q", "r:R"],
    from: 0,
    intent: "range",
    expected: ["a:Q", "r:R"],
  },
};

describe("deleting a break beside pending joins", () => {
  for (const [name, { specs, from, intent, expected }] of Object.entries(cases))
    for (const mode of ["direct", "tracked"] as const)
      test(`${name}, ${mode}`, () => {
        const document = documentOf(specs);
        expect(
          accepted(document, intentsAt(document, from)[intent], MODES[mode]),
        ).toEqual([...expected]);
      });

  test.each([
    [["w:P", "bD:Q", ":R"], 0, "w:P"],
    [["w:P", "D:Q", ":R"], 0, "w:P"],
    [["aD:P", ":O", "bD:Q", "D:S", ":R"], 1, "a:O"],
    [[":P", "bD:Q", ":R"], 0, ":R"],
  ] satisfies [Spec[], number, string][])(
    "a tracked deletion in %j that retires the break before its container's pending end keeps the words' properties",
    (specs, from, expected) => {
      const document = documentOf(specs);
      expect(
        accepted(document, intentsAt(document, from).deleteNext, MODES.tracked),
      ).toEqual([expected]);
    },
  );

  test("a tracked deletion before a pending end with a trailing bookmark keeps the words' properties", () => {
    const document = documentOf(["w:P", "bD:Q", ":R"]);
    const [first] = paragraphsOf(document);
    if (first === undefined)
      panic("The bookmark starts in the first paragraph.");
    first.content.unshift({ type: "bookmarkStart", id: 11, name: "Clause" });
    document.package.document.content.push({ type: "bookmarkEnd", id: 11 });
    expect(
      accepted(document, intentsAt(document, 0).deleteNext, MODES.tracked),
    ).toEqual(["w:P"]);
  });

  test("a last paragraph a pending mark already joins keeps its own properties, deleted either way", () => {
    const document = documentOf(["aD:P", "D:Q", "r:R"]);
    const { deleteNext } = intentsAt(document, 1);
    expect(accepted(document, deleteNext, MODES.tracked)).toEqual(["a:R"]);
    expect(
      accepted(
        documentOf(["aD:P", "r:R"]),
        intentsAt(documentOf(["aD:P", "r:R"]), 0).deleteNext,
        MODES.direct,
      ),
    ).toEqual(["a:R"]);
  });
});
