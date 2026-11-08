import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { Document, ParagraphContent } from "../types/document";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";

type Lock = "contentLocked" | "sdtLocked" | "sdtContentLocked";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

/** A rich text control holding "Clause" inside one paragraph. */
const controlDocument = (lock?: Lock): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [
            run("Before "),
            {
              type: "inlineSdt",
              properties: {
                sdtType: "richText",
                tag: "clause",
                id: 7,
                ...(lock === undefined ? {} : { lock }),
              },
              content: [run("Clause")],
            },
            run(" after"),
          ],
        },
        {
          type: "paragraph",
          paraId: "20000000",
          content: [run("Next")],
        },
      ],
    },
  },
});

const session = async (
  lock?: Lock,
  mode: "editing" | "suggesting" = "editing",
) => {
  const created = createCanonicalSession(
    await parseDocx(
      (
        await prepareCanonicalDocxInput(await createDocx(controlDocument(lock)))
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();
  if (mode === "suggesting")
    created.setMode({ type: "suggesting", author: "Reviewer" });
  return created;
};

const stateFor = (current: CanonicalSession) =>
  EditorState.create({ schema, doc: current.projection.doc });
const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  expect(next.doc.eq(commit.transaction.doc)).toBe(true);
  return next;
};
/** The position `offset` characters into the text node `text`. */
const at = (state: EditorState, text: string, offset = text.length) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === text) found = position + offset;
    return true;
  });
  return found ?? panic(`No text ${text}`);
};
/** The positions just outside and just inside the control's edges. */
const edges = (state: EditorState) => {
  let found: { open: number; close: number } | undefined;
  state.doc.descendants((node, position) => {
    if (node.type.name === "sdt")
      found = { open: position, close: position + node.nodeSize };
    return found === undefined;
  });
  return found ?? panic("No control");
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

/** The control's text and the paragraph's text outside it, from the model. */
const parts = (current: CanonicalSession) => {
  const paragraph = current.document.package.document.content[0];
  if (paragraph?.type !== "paragraph") return panic("No paragraph");
  const textOf = (items: readonly ParagraphContent[]): string =>
    items
      .map((item) => {
        if (item.type === "run")
          return item.content
            .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
            .join("");
        if (item.type === "inlineSdt") return `[${textOf(item.content)}]`;
        if ("content" in item && Array.isArray(item.content))
          return textOf(item.content as ParagraphContent[]);
        return "";
      })
      .join("");
  return textOf(paragraph.content);
};

describe("inline content controls in the canonical session", () => {
  test("open, and typing inside saves into the control", async () => {
    const current = await session();
    expect(parts(current)).toBe("Before [Clause] after");
    let state = stateFor(current);
    const inside = at(state, "Clause", 3);
    state = select(state, inside);
    state = accept(
      state,
      current
        .prepareReplace(state, { from: inside, to: inside, text: "X" })
        .unwrap(),
    );
    expect(parts(current)).toBe("Before [ClaXuse] after");
    const saved = await serializeCanonicalSave({
      snapshot: current.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    });
    const reopened = createCanonicalSession(
      await parseDocx(saved.buffer, {
        preloadFonts: false,
        detectVariables: false,
      }),
    ).unwrap();
    expect(parts(reopened)).toBe("Before [ClaXuse] after");
  });

  for (const mode of ["editing", "suggesting"] as const)
    test(`typing at each edge of the control in ${mode} mode lands on the side the caret is`, async () => {
      for (const edge of ["open", "close"] as const)
        for (const side of ["outside", "inside"] as const) {
          const current = await session(undefined, mode);
          let state = stateFor(current);
          const { open, close } = edges(state);
          const position = {
            "open:outside": open,
            "open:inside": open + 1,
            "close:inside": close - 1,
            "close:outside": close,
          }[`${edge}:${side}`];
          state = select(state, position);
          state = accept(
            state,
            current
              .prepareReplace(state, {
                from: position,
                to: position,
                text: "Z",
              })
              .unwrap(),
          );
          expect(state.doc.eq(current.projection.doc)).toBe(true);
          // Inside an edge the text joins the control; outside it stays out.
          expect(parts(current)).toBe(
            {
              "open:outside": "Before Z[Clause] after",
              "open:inside": "Before [ZClause] after",
              "close:inside": "Before [ClauseZ] after",
              "close:outside": "Before [Clause]Z after",
            }[`${edge}:${side}`],
          );
        }
    });

  test("a deletion across an edge keeps the control", async () => {
    const current = await session();
    let state = stateFor(current);
    const from = at(state, "Before ", 3);
    const to = at(state, "Clause", 3);
    state = select(state, from, to);
    state = accept(
      state,
      current.prepareReplace(state, { from, to, text: "" }).unwrap(),
    );
    expect(parts(current)).toBe("Bef[use] after");
  });

  test("Enter inside the control splits it, the second half under a fresh id, and undo is exact", async () => {
    const current = await session();
    const before = current.document;
    let state = stateFor(current);
    state = select(state, at(state, "Clause", 3));
    state = accept(state, current.prepareSplit(state).unwrap());
    const ids = current.document.package.document.content.flatMap((block) =>
      block.type === "paragraph"
        ? block.content.flatMap((item) =>
            item.type === "inlineSdt" ? [item.properties.id] : [],
          )
        : [],
    );
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe(7);
    expect(new Set(ids).size).toBe(2);
    state = accept(state, current.prepareUndo(state).unwrap());
    expect(current.document).toEqual(before);
  });

  /** Every tracked change id in the session's main story. */
  const revisionIds = (current: CanonicalSession): number[] => {
    const ids: number[] = [];
    const visit = (items: readonly ParagraphContent[]) => {
      for (const item of items) {
        if (
          item.type === "insertion" ||
          item.type === "deletion" ||
          item.type === "moveFrom" ||
          item.type === "moveTo"
        )
          ids.push(item.info.id);
        if ("content" in item && Array.isArray(item.content))
          visit(item.content as ParagraphContent[]);
      }
    };
    for (const block of current.document.package.document.content)
      if (block.type === "paragraph") visit(block.content);
    return [...new Set(ids)];
  };

  test("a suggestion over the control's whole text refuses and changes nothing", async () => {
    const current = await session(undefined, "suggesting");
    const before = current.document;
    let state = stateFor(current);
    const { open, close } = edges(state);
    state = select(state, open + 1, close - 1);
    expect(
      current
        .prepareReplace(state, { from: open + 1, to: close - 1, text: "New" })
        .isErr(),
    ).toBe(true);
    expect(current.document).toBe(before);
  });

  test("replacing all of the control's text in editing mode keeps the new text in it", async () => {
    const current = await session(undefined, "editing");
    let state = stateFor(current);
    const { open, close } = edges(state);
    state = select(state, open + 1, close - 1);
    accept(
      state,
      current
        .prepareReplace(state, { from: open + 1, to: close - 1, text: "New" })
        .unwrap(),
    );
    expect(parts(current)).toBe("Before [New] after");
  });

  test(`replacing part of the control's text keeps the new text in it, and a suggestion resolves to the edit or the source`, async () => {
    const replaced = async (mode: "editing" | "suggesting") => {
      const current = await session(undefined, mode);
      let state = stateFor(current);
      const from = at(state, "Clause", 1);
      const to = at(state, "Clause", 4);
      state = select(state, from, to);
      state = accept(
        state,
        current.prepareReplace(state, { from, to, text: "New" }).unwrap(),
      );
      return { current, state };
    };
    const edited = await replaced("editing");
    const expected = parts(edited.current);
    expect(expected.split("[")).toHaveLength(2);
    expect(expected).toContain("New");
    const source = parts(await session(undefined, "editing"));
    for (const resolution of ["accept", "reject"] as const) {
      const suggested = await replaced("suggesting");
      expect(parts(suggested.current).split("[")).toHaveLength(2);
      accept(
        suggested.state,
        suggested.current
          .prepareResolve(suggested.state, {
            revisionIds: revisionIds(suggested.current),
            resolution,
          })
          .unwrap(),
      );
      expect(parts(suggested.current)).toBe(
        resolution === "accept" ? expected : source,
      );
    }
  });

  for (const mode of ["editing", "suggesting"] as const)
    test(`deleting text inside the control in ${mode} mode keeps the control`, async () => {
      const current = await session(undefined, mode);
      let state = stateFor(current);
      const from = at(state, "Clause", 1);
      const to = at(state, "Clause", 4);
      state = select(state, from, to);
      state = accept(
        state,
        current.prepareReplace(state, { from, to, text: "" }).unwrap(),
      );
      expect(parts(current).startsWith("Before [C")).toBe(true);
      expect(parts(current).split("[")).toHaveLength(2);
    });

  for (const lock of ["contentLocked", "sdtContentLocked"] as const)
    test(`${lock}: typing inside refuses as the document's rule`, async () => {
      const current = await session(lock);
      const before = current.document;
      let state = stateFor(current);
      const inside = at(state, "Clause", 3);
      state = select(state, inside);
      const typed = current.prepareReplace(state, {
        from: inside,
        to: inside,
        text: "X",
      });
      expect(typed.isErr()).toBe(true);
      if (typed.isErr())
        expect(typed.error.ground).toEqual({ rule: "contentControlLock" });
      expect(current.document).toBe(before);
    });

  test("sdtLocked: deleting the whole control refuses, editing inside works", async () => {
    const current = await session("sdtLocked");
    let state = stateFor(current);
    const from = at(state, "Before ", 3);
    const to = at(state, " after", 2);
    state = select(state, from, to);
    const removed = current.prepareReplace(state, { from, to, text: "" });
    expect(removed.isErr()).toBe(true);
    if (removed.isErr())
      expect(removed.error.ground).toEqual({ rule: "contentControlLock" });
    const inside = at(state, "Clause", 3);
    state = select(state, inside);
    state = accept(
      state,
      current
        .prepareReplace(state, { from: inside, to: inside, text: "X" })
        .unwrap(),
    );
    expect(parts(current)).toBe("Before [ClaXuse] after");
  });
});
