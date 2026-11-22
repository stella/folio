/**
 * One range per move-range id: an inserted copy of a move range boundary that
 * stands in the document is left out, in every mode, and the insertion is
 * undone exactly. A boundary of a range the document does not hold, or one
 * completing a range, goes in as it is.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  ParagraphContent,
  Run,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionStamp,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const toStart = (id: number): ParagraphContent => ({
  type: "moveToRangeStart",
  id,
  name: `move${id}`,
});
const toEnd = (id: number): ParagraphContent => ({
  type: "moveToRangeEnd",
  id,
});
const fromStart = (id: number): ParagraphContent => ({
  type: "moveFromRangeStart",
  id,
  name: `move${id}`,
});
const fromEnd = (id: number): ParagraphContent => ({
  type: "moveFromRangeEnd",
  id,
});

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content },
        { type: "paragraph", paraId: "00000002", content: [run("tail")] },
      ],
    },
  },
});

const boundariesIn = (document: Document): string[] => {
  const out: string[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const type: unknown = Reflect.get(value, "type");
    if (typeof type === "string" && /^move(?:From|To)Range/u.test(type))
      out.push(`${type}:${String(Reflect.get(value, "id"))}`);
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return out;
};

const textOf = (document: Document): string => {
  let text = "";
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    if (Reflect.get(value, "type") === "text")
      text += String(Reflect.get(value, "text"));
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return text;
};

const MODES = {
  direct: undefined,
  tracked: { id: 90, author: "R", date: "2026-01-01T00:00:00Z" },
  suggested: {
    id: 90,
    author: "R",
    date: "2026-01-01T00:00:00Z",
    suggestionId: "s1",
  },
} as const satisfies Record<string, RevisionStamp | undefined>;

const insert = (
  content: ParagraphContent[],
  revision: RevisionStamp | undefined,
): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
  at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 4 },
  slice: { content, openStart: 0, openEnd: 0 },
  ...(revision === undefined
    ? {}
    : { revision, newIds: { revision: [91, 92, 93] } }),
});

/** `op` applied, undone after a JSON round trip and redone. */
const applied = (document: Document, op: DocumentOp) => {
  const before = structuredClone(document);
  const result = applyDocumentOp(document, JSON.parse(JSON.stringify(op)));
  if (result.isErr()) throw result.error;
  const undone = applyDocumentOps(
    result.value.document,
    JSON.parse(JSON.stringify(result.value.inverse)),
  );
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toEqual(before);
  const redone = applyDocumentOps(
    undone.value.document,
    JSON.parse(JSON.stringify(undone.value.inverse)),
  );
  if (redone.isErr()) throw redone.error;
  expect(redone.value.document).toEqual(result.value.document);
  return result.value;
};

for (const [mode, revision] of Object.entries(MODES))
  describe(mode, () => {
    test("a copy of a whole standing range is its text alone", () => {
      for (const [open, close] of [
        [toStart, toEnd],
        [fromStart, fromEnd],
      ] as const) {
        const document = documentOf([open(1), run("ab"), close(1), run("cd")]);
        const result = applied(
          document,
          insert([open(1), run("ab"), close(1)], revision),
        );
        expect(boundariesIn(result.document)).toEqual(boundariesIn(document));
        expect(textOf(result.document)).toBe("abcdtailab");
        if (revision === undefined) return;
        const rejected = applyDocumentOp(result.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: result.revisions,
          decision: REVISION_DECISIONS.REJECT,
        });
        if (rejected.isErr()) throw rejected.error;
        expect(rejected.value.document).toEqual(document);
      }
    });

    test("half a standing range is its text alone", () => {
      const document = documentOf([toStart(1), run("ab"), toEnd(1)]);
      const result = applied(document, insert([run("b"), toEnd(1)], revision));
      expect(boundariesIn(result.document)).toEqual(boundariesIn(document));
      expect(textOf(result.document)).toBe("abtailb");
    });

    test("a slice holding only a repeated boundary inserts nothing", () => {
      const document = documentOf([toStart(1), run("ab"), toEnd(1)]);
      expect(applied(document, insert([toEnd(1)], revision)).document).toEqual(
        document,
      );
    });

    test("a range the document does not hold goes in as it is", () => {
      const document = documentOf([run("ab")]);
      const result = applied(
        document,
        insert([toStart(1), run("x"), toEnd(1)], revision),
      );
      expect(boundariesIn(result.document)).toEqual([
        "moveToRangeStart:1",
        "moveToRangeEnd:1",
      ]);
    });

    test("a boundary completing a range goes in as it is", () => {
      const document = documentOf([toStart(1), run("ab")]);
      const result = applied(document, insert([run("x"), toEnd(1)], revision));
      expect(boundariesIn(result.document)).toEqual([
        "moveToRangeStart:1",
        "moveToRangeEnd:1",
      ]);
    });

    test("a whole range beside a standing half is its text alone", () => {
      const document = documentOf([run("ab"), toEnd(1)]);
      const result = applied(
        document,
        insert([toStart(1), run("x"), toEnd(1)], revision),
      );
      expect(boundariesIn(result.document)).toEqual(["moveToRangeEnd:1"]);
      expect(textOf(result.document)).toBe("abtailx");
    });

    test("the slice's first range is the range, where it repeats one", () => {
      const document = documentOf([run("ab")]);
      const result = applied(
        document,
        insert(
          [toStart(1), run("x"), toEnd(1), toStart(1), run("y"), toEnd(1)],
          revision,
        ),
      );
      expect(boundariesIn(result.document)).toEqual([
        "moveToRangeStart:1",
        "moveToRangeEnd:1",
      ]);
      expect(textOf(result.document)).toBe("abtailxy");
    });

    test("a range standing between table rows' cells is standing", () => {
      const table: Table = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000009",
                    content: [toStart(1), run("c"), toEnd(1)],
                  },
                ],
              },
            ],
          },
        ],
      };
      const document: Document = {
        package: {
          document: {
            content: [
              table,
              { type: "paragraph", paraId: "00000002", content: [run("tail")] },
            ],
          },
        },
      };
      const result = applied(
        document,
        insert([toStart(1), run("c"), toEnd(1)], revision),
      );
      expect(boundariesIn(result.document)).toEqual(boundariesIn(document));
    });

    test("a range in a pasted link is its text alone", () => {
      const document = documentOf([toStart(1), run("ab"), toEnd(1)]);
      const result = applied(
        document,
        insert(
          [
            {
              type: "hyperlink",
              anchor: "here",
              children: [toStart(1), run("ab"), toEnd(1)],
            },
          ],
          revision,
        ),
      );
      expect(boundariesIn(result.document)).toEqual(boundariesIn(document));
    });
  });

describe("a pasted fragment", () => {
  const at = (offset: number) => ({
    story: OP_STORIES.MAIN,
    blockId: "00000002",
    offset,
  });
  const source = (): Document =>
    documentOf([toStart(1), run("ab"), toEnd(1), run("cd")]);
  const pasteModes = {
    editing: { type: "editing" },
    suggesting: {
      type: "suggesting",
      revision: MODES.tracked,
      newIds: { revision: [91, 92, 93, 94, 95] },
    },
  } as const satisfies Record<string, EditorIntentMode>;
  for (const [name, mode] of Object.entries(pasteModes)) {
    test(`${name}: copies of one or several paragraphs are their text alone`, () => {
      for (const paragraphs of [
        [[toStart(1), run("ab"), toEnd(1)]],
        [
          [toStart(1), run("a")],
          [run("b"), toEnd(1)],
        ],
        [[run("a"), toStart(1)], [toEnd(1)], [run("b")]],
      ]) {
        const document = source();
        const compiled = compileEditorIntent(document, {
          intent: {
            type: "replaceFragment",
            from: at(4),
            to: at(4),
            openStart: 1,
            openEnd: 1,
            paragraphs: paragraphs.map((content, index) => ({
              type: "paragraph",
              paraId: `0000001${index}`,
              content,
            })),
          },
          mode,
        });
        if (compiled.isErr()) throw compiled.error;
        const after = applyDocumentOps(document, compiled.value.ops);
        if (after.isErr()) throw after.error;
        expect(boundariesIn(after.value.document)).toEqual(
          boundariesIn(document),
        );
        expect(textOf(after.value.document)).toBe("abcdtailab");
        const undone = applyDocumentOps(
          after.value.document,
          JSON.parse(JSON.stringify(after.value.inverse)),
        );
        if (undone.isErr()) throw undone.error;
        expect(undone.value.document).toEqual(source());
      }
    });
  }

  test("pasted over its own source, the range is the range", () => {
    const document = source();
    const here = (offset: number) => ({ ...at(offset), blockId: "00000001" });
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "replaceFragment",
        from: { ...here(0), zeroWidthBefore: 0 },
        to: { ...here(2), zeroWidthBefore: 1 },
        openStart: 1,
        openEnd: 1,
        paragraphs: [
          {
            type: "paragraph",
            paraId: "00000010",
            content: [toStart(1), run("xy"), toEnd(1)],
          },
        ],
      },
      mode: { type: "editing" },
    });
    if (compiled.isErr()) throw compiled.error;
    const after = applyDocumentOps(document, compiled.value.ops);
    if (after.isErr()) throw after.error;
    expect(
      boundariesIn(after.value.document).map((key) => key.split(":").at(0)),
    ).toEqual(["moveToRangeStart", "moveToRangeEnd"]);
    expect(textOf(after.value.document)).toBe("xycdtail");
  });
});

test("undoing a deletion puts back every boundary its slice repeats", () => {
  applied(
    documentOf([
      run("ab"),
      toStart(1),
      run("cd"),
      toEnd(1),
      toStart(1),
      run("ef"),
      toEnd(1),
      run("gh"),
    ]),
    {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 7 },
    },
  );
});

test("a bookmark copy stays a copy", () => {
  const document = documentOf([
    { type: "bookmarkStart", id: 1, name: "_Ref1" },
    run("ab"),
    { type: "bookmarkEnd", id: 1 },
  ]);
  const result = applied(
    document,
    insert(
      [
        { type: "bookmarkStart", id: 1, name: "_Ref1" },
        run("ab"),
        { type: "bookmarkEnd", id: 1 },
      ],
      MODES.tracked,
    ),
  );
  expect(JSON.stringify(result.document).match(/bookmarkStart/gu)).toHaveLength(
    2,
  );
});
