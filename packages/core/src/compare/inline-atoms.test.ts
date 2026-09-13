import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import { matchInlineAtoms } from "./inline-atoms";

const field = () =>
  schema.nodes["field"]!.create({
    fieldType: "NUMPAGES",
    instruction: " NUMPAGES ",
    displayText: "1",
    fieldKind: "simple",
  });

const wordsField = () =>
  schema.nodes["field"]!.create({
    fieldType: "NUMWORDS",
    instruction: " NUMWORDS ",
    displayText: "2",
    fieldKind: "simple",
  });

const documentWith = (content: readonly PMNode[]) =>
  schema.node("doc", null, [schema.node("paragraph", null, content)]);

const resolve = ({ state, mode }: { state: EditorState; mode: "accept" | "reject" }): EditorState => {
  let resolved = state;
  const command = mode === "accept" ? acceptAllChanges() : rejectAllChanges();
  command(state, (transaction) => {
    resolved = state.apply(transaction);
  });
  return resolved;
};

describe("matchInlineAtoms", () => {
  test("tracks a field restored into replacement text through accept and reject", () => {
    const target = documentWith([schema.text("Pages: "), field()]);
    const state = EditorState.create({ schema, doc: documentWith([schema.text("Pages: ")]) });
    const result = matchInlineAtoms({
      state,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);
    expect(resolve({ state: reviewed, mode: "accept" }).doc.eq(target)).toBe(true);
    expect(resolve({ state: reviewed, mode: "reject" }).doc.textContent).toBe("Pages: ");
    const repeated = matchInlineAtoms({
      state: reviewed,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: result.nextRevisionId, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });
    expect(repeated).toMatchObject({ status: "matched", rangeCount: 0 });
  });

  test("maps an accepted field boundary back through a deleted paragraph mark", () => {
    const target = documentWith([schema.text("Pages: "), field()]);
    const source = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          pPrMark: {
            kind: "del",
            info: { id: 8, author: "Compare", date: "2026-09-13T00:00:00.000Z" },
          },
        },
        [schema.text("Pages:")],
      ),
      schema.node("paragraph", null, [schema.text(" ")]),
    ]);
    const state = EditorState.create({ schema, doc: source });
    expect(resolve({ state, mode: "accept" }).doc.textContent).toBe("Pages: ");
    const result = matchInlineAtoms({
      state,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);
    expect(resolve({ state: reviewed, mode: "accept" }).doc.eq(target)).toBe(true);
    expect(resolve({ state: reviewed, mode: "reject" }).doc.eq(resolve({ state, mode: "reject" }).doc)).toBe(true);
  });

  test("preserves the target order for multiple atoms at one text boundary", () => {
    const target = documentWith([schema.text("Counts: "), field(), wordsField()]);
    const state = EditorState.create({ schema, doc: documentWith([schema.text("Counts: ")]) });
    const result = matchInlineAtoms({
      state,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    expect(resolve({ state: state.apply(result.transaction), mode: "accept" }).doc.eq(target)).toBe(true);
  });

  test("tracks a zero-width page break without changing surrounding text", () => {
    const target = documentWith([schema.text("Before"), schema.node("pageBreakRun"), schema.text("After")]);
    const state = EditorState.create({ schema, doc: documentWith([schema.text("BeforeAfter")]) });
    const result = matchInlineAtoms({
      state,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);
    expect(resolve({ state: reviewed, mode: "accept" }).doc.eq(target)).toBe(true);
    expect(resolve({ state: reviewed, mode: "reject" }).doc.textContent).toBe("BeforeAfter");
  });

  test("uses the stable paragraph identity when review resolution deletes its position", () => {
    const attrs = {
      paraId: "00A0B0C0",
      pPrMark: {
        kind: "ins",
        info: { id: 8, author: "Compare", date: "2026-09-13T00:00:00.000Z" },
      },
    };
    const source = schema.node("doc", null, [schema.node("paragraph", attrs)]);
    const target = schema.node("doc", null, [
      schema.node("paragraph", { ...attrs, pPrMark: null }, [schema.node("pageBreakRun")]),
    ]);
    const state = EditorState.create({ schema, doc: source });
    const result = matchInlineAtoms({
      state,
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);
    expect(resolve({ state: reviewed, mode: "accept" }).doc.eq(target)).toBe(true);
    const rejected = resolve({ state: reviewed, mode: "reject" }).doc;
    expect(rejected.firstChild?.childCount).toBe(0);
    expect(rejected.firstChild?.attrs["pPrMark"]).toBeNull();
  });

  test("refuses a deleted mapping when the paragraph identity is duplicated", () => {
    const attrs = {
      paraId: "00A0B0C0",
      pPrMark: {
        kind: "ins",
        info: { id: 8, author: "Compare", date: "2026-09-13T00:00:00.000Z" },
      },
    };
    const source = schema.node("doc", null, [
      schema.node("paragraph", attrs),
      schema.node("paragraph", {
        ...attrs,
        pPrMark: {
          kind: "del",
          info: { id: 9, author: "Compare", date: "2026-09-13T00:00:00.000Z" },
        },
      }),
    ]);
    const target = schema.node("doc", null, [
      schema.node("paragraph", { ...attrs, pPrMark: null }, [schema.node("pageBreakRun")]),
    ]);
    const result = matchInlineAtoms({
      state: EditorState.create({ schema, doc: source }),
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 10,
    });

    expect(result.status).toBe("unalignable");
  });

  test("reports the range budget before creating an atom revision", () => {
    const target = documentWith([schema.text("Pages: "), field()]);
    const result = matchInlineAtoms({
      state: EditorState.create({ schema, doc: documentWith([schema.text("Pages: ")]) }),
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 0,
    });

    expect(result.status).toBe("budget-exceeded");
  });
});
