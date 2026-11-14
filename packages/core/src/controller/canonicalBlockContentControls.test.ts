import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

type Lock = "contentLocked" | "sdtLocked" | "sdtContentLocked";

/** A clause inside a rich text block control, between two plain paragraphs. */
const controlDocument = (lock?: Lock, pending = false): Document => ({
  package: {
    document: {
      content: [
        paragraph("10000000", "Before"),
        {
          type: "blockSdt",
          properties: {
            sdtType: "richText",
            tag: "clause",
            alias: "Clause",
            id: 7,
            ...(lock === undefined ? {} : { lock }),
          },
          content: [
            pending
              ? {
                  ...paragraph("20000000", "First clause"),
                  content: [
                    ...paragraph("20000000", "First clause").content,
                    {
                      type: "insertion",
                      info: {
                        id: 41,
                        author: "Earlier",
                        date: "2026-11-01T10:00:00Z",
                      },
                      content: [
                        {
                          type: "run",
                          content: [{ type: "text", text: " added" }],
                        },
                      ],
                    },
                  ],
                }
              : paragraph("20000000", "First clause"),
            paragraph("21000000", "Second clause"),
          ],
        },
        pending
          ? {
              ...paragraph("30000000", "After"),
              content: [
                ...paragraph("30000000", "After").content,
                {
                  type: "insertion",
                  info: {
                    id: 42,
                    author: "Earlier",
                    date: "2026-11-01T10:00:00Z",
                  },
                  content: [
                    { type: "run", content: [{ type: "text", text: " too" }] },
                  ],
                },
              ],
            }
          : paragraph("30000000", "After"),
      ],
    },
  },
});

const session = async (lock?: Lock, pending = false) =>
  createCanonicalSession(
    await parseDocx(
      (
        await prepareCanonicalDocxInput(
          await createDocx(controlDocument(lock, pending)),
        )
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();

const stateFor = (current: CanonicalSession) =>
  EditorState.create({ schema, doc: current.projection.doc });
const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};
const at = (state: EditorState, text: string, offset = text.length) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === text) found = position + offset;
    return true;
  });
  return found ?? panic(`No text ${text}`);
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
const control = (document: Document) => {
  const block = document.package.document.content.find(
    (item): item is Extract<BlockContent, { type: "blockSdt" }> =>
      item.type === "blockSdt",
  );
  return block ?? panic("The control is gone.");
};
const texts = (blocks: readonly BlockContent[]) =>
  blocks.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) =>
            item.type === "run"
              ? item.content.flatMap((leaf) =>
                  leaf.type === "text" ? [leaf.text] : [],
                )
              : [],
          )
          .join("")
      : block.type,
  );

describe("block content controls in the canonical session", () => {
  test("open, and typing inside a control stays inside it through save", async () => {
    const current = await session();
    let state = stateFor(current);
    const end = at(state, "First clause");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "s" }).unwrap(),
    );
    expect(texts(control(current.document).content)).toEqual([
      "First clauses",
      "Second clause",
    ]);
    const saved = await serializeCanonicalSave({
      snapshot: current.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    });
    const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
    const kept = control(reopened);
    expect(texts(kept.content)).toEqual(["First clauses", "Second clause"]);
    expect(kept.properties.tag).toBe("clause");
    expect(kept.properties.alias).toBe("Clause");
  });

  test("Enter inside a control adds a paragraph inside it, and undo removes it", async () => {
    const current = await session();
    let state = stateFor(current);
    const before = current.document;
    state = select(state, at(state, "First clause", 5));
    state = accept(state, current.prepareSplit(state).unwrap());
    expect(texts(control(current.document).content)).toEqual([
      "First",
      " clause",
      "Second clause",
    ]);
    state = accept(state, current.prepareUndo(state).unwrap());
    expect(current.document).toEqual(before);
  });

  test("a join never crosses a control's edge", async () => {
    const current = await session();
    const before = current.document;
    for (const [text, offset, direction] of [
      ["First clause", 0, "backward"],
      ["Second clause", 13, "forward"],
      ["After", 0, "backward"],
      ["Before", 6, "forward"],
    ] as const) {
      const state = select(
        stateFor(current),
        at(stateFor(current), text, offset),
      );
      const joined = current.prepareJoin(state, direction);
      expect([text, joined.isErr()]).toEqual([text, true]);
    }
    expect(current.document).toBe(before);
  });

  test("typing over a selection that crosses a control's edge keeps the control and joins nothing across it", async () => {
    const current = await session();
    const before = current.document;
    let state = stateFor(current);
    const start = at(state, "Before", 2);
    const end = at(state, "First clause", 3);
    state = select(state, start, end);
    state = accept(
      state,
      current
        .prepareReplace(state, { from: start, to: end, text: "x" })
        .unwrap(),
    );
    const blocks = current.document.package.document.content;
    expect(blocks.map(({ type }) => type)).toEqual([
      "paragraph",
      "blockSdt",
      "paragraph",
    ]);
    expect(texts([blocks[0] ?? panic("No first block.")])).toEqual(["Bex"]);
    expect(texts(control(current.document).content)).toEqual([
      "st clause",
      "Second clause",
    ]);
    state = accept(state, current.prepareUndo(state).unwrap());
    expect(current.document).toEqual(before);
  });

  test("typing inside a control is tracked in suggesting mode and rejects to the source", async () => {
    const current = await session();
    current.setMode({ type: "suggesting", author: "Reviewer" });
    let state = stateFor(current);
    const end = at(state, "Second clause");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "!" }).unwrap(),
    );
    const record = JSON.stringify(control(current.document).content);
    expect(record).toContain('"insertion"');
  });

  test("a list paragraph inside a control opens, edits and keeps its numbering", async () => {
    const document = controlDocument();
    const block = document.package.document.content[1];
    if (block?.type !== "blockSdt") return panic("No control.");
    const first = block.content[0];
    if (first?.type !== "paragraph") return panic("No clause.");
    block.content[0] = {
      ...first,
      formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
    };
    const current = createCanonicalSession(document).unwrap();
    let state = stateFor(current);
    const end = at(state, "First clause");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "s" }).unwrap(),
    );
    const edited = control(current.document).content[0];
    if (edited?.type !== "paragraph") return panic("No clause.");
    expect(texts([edited])).toEqual(["First clauses"]);
    expect(edited.formatting?.numPr).toEqual({
      kind: "reference",
      numId: 1,
      ilvl: 0,
    });
  });

  describe("locks", () => {
    for (const lock of ["contentLocked", "sdtContentLocked"] as const)
      test(`${lock}: typing inside refuses as the document's rule and changes nothing`, async () => {
        const current = await session(lock);
        const before = current.document;
        for (const mode of ["editing", "suggesting"] as const) {
          current.setMode(
            mode === "editing"
              ? { type: "editing" }
              : { type: "suggesting", author: "Reviewer" },
          );
          let state = stateFor(current);
          const end = at(state, "First clause");
          state = select(state, end);
          const refused = current.prepareReplace(state, {
            from: end,
            to: end,
            text: "s",
          });
          expect(refused.isErr()).toBe(true);
          if (refused.isErr())
            expect(refused.error.ground).toEqual({
              rule: "contentControlLock",
            });
          // A selection reaching in from outside meets the same rule.
          const start = at(state, "Before", 2);
          state = select(state, start, end);
          expect(
            current
              .prepareReplace(state, { from: start, to: end, text: "x" })
              .isErr(),
          ).toBe(true);
          expect(current.document).toBe(before);
        }
      });

    test("sdtLocked: the control's content stays editable", async () => {
      const current = await session("sdtLocked");
      let state = stateFor(current);
      const end = at(state, "First clause");
      state = select(state, end);
      state = accept(
        state,
        current
          .prepareReplace(state, { from: end, to: end, text: "s" })
          .unwrap(),
      );
      expect(texts(control(current.document).content)).toEqual([
        "First clauses",
        "Second clause",
      ]);
    });

    test("accepting holds back a revision inside a locked control and resolves the rest", async () => {
      const current = await session("contentLocked", true);
      const state = stateFor(current);
      const prepared = current
        .prepareResolve(state, { revisionIds: [41, 42], resolution: "accept" })
        .unwrap();
      expect(prepared.lockHeld).toEqual([41]);
      expect(prepared.publish().isOk()).toBe(true);
      const record = JSON.stringify(current.document.package.document.content);
      // The held insertion is still pending; the one outside was accepted.
      expect(record).toContain('"id":41');
      expect(record).not.toContain('"id":42');
      // Resolving only the held revision refuses outright, as the rule.
      const alone = current.prepareResolve(stateFor(current), {
        revisionIds: [41],
        resolution: "reject",
      });
      expect(alone.isErr()).toBe(true);
      if (alone.isErr())
        expect(alone.error.ground).toEqual({ rule: "contentControlLock" });
    });

    test("a public batch skips the operation a lock refuses and applies the rest", async () => {
      const current = await session("contentLocked");
      let state = stateFor(current);
      const executor = new CanonicalPublicOperations({
        session: current,
        getState: () => state,
        publish: (commit) => {
          const published = publishCanonicalProjection({
            state,
            commit,
            session: current,
          });
          if (published.isErr()) return false;
          state = published.value.state;
          return true;
        },
      });
      const result = executor.apply({
        snapshot: createFolioAIEditSnapshot(state.doc),
        undefinedReferences: "refuse",
        batch: {
          version: 1,
          mode: "direct",
          operations: [
            {
              id: "inside",
              type: "replaceInBlock",
              blockId: "20000000",
              find: "First",
              replace: "Only",
            },
            {
              id: "outside",
              type: "replaceInBlock",
              blockId: "30000000",
              find: "After",
              replace: "Later",
            },
          ],
        },
      });
      expect(result.applied.map(({ id }) => id)).toEqual(["outside"]);
      expect(result.skipped.map(({ id, reason }) => [id, reason])).toEqual([
        ["inside", "contentControlLock"],
      ]);
      expect(texts(control(current.document).content)).toEqual([
        "First clause",
        "Second clause",
      ]);
      expect(texts(current.document.package.document.content)).toContain(
        "Later",
      );
    });
  });
});
