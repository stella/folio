import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import type { ParagraphContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";
import { matchInlineAtoms, sameInlineAtoms } from "./inline-atoms";

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

const comparisonInsertedText = (text: string) =>
  schema.text(text, [
    schema.marks["insertion"]!.create({
      revisionId: 20,
      author: "Compare",
      date: "2026-09-13T00:00:00.000Z",
    }),
  ]);

const docxWith = async (content: readonly ParagraphContent[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "12345678", content: [...content] },
  ];
  return await createDocx(document);
};

const resolve = ({
  state,
  mode,
}: {
  state: EditorState;
  mode: "accept" | "reject";
}): EditorState => {
  let resolved = state;
  const command = mode === "accept" ? acceptAllChanges() : rejectAllChanges();
  command(state, (transaction) => {
    resolved = state.apply(transaction);
  });
  return resolved;
};

describe("matchInlineAtoms", () => {
  test("replaces a materialized field result with its serialized carrier", async () => {
    const base = await docxWith([{ type: "run", content: [{ type: "text", text: "Count: 12" }] }]);
    const target = await docxWith([
      { type: "run", content: [{ type: "text", text: "Count: " }] },
      {
        type: "simpleField",
        fieldType: "NUMWORDS",
        instruction: " NUMWORDS ",
        content: [{ type: "run", content: [{ type: "text", text: "12" }] }],
      },
    ]);

    const compared = await compareDocx(base, target, {
      author: "Compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    });

    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
    const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    accepting.acceptAll();
    expect(accepting.snapshot().blocks.at(0)?.text).toBe("Count: 12");
    const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    rejecting.rejectAll();
    expect(rejecting.snapshot().blocks.at(0)?.text).toBe("Count: 12");
  });

  test("replaces a field carrier with retained literal result text", async () => {
    const base = await docxWith([
      { type: "run", content: [{ type: "text", text: "Count: " }] },
      {
        type: "simpleField",
        fieldType: "NUMWORDS",
        instruction: " NUMWORDS ",
        content: [{ type: "run", content: [{ type: "text", text: "12" }] }],
      },
    ]);
    const target = await docxWith([
      { type: "run", content: [{ type: "text", text: "Count: 12" }] },
    ]);

    const compared = await compareDocx(base, target, {
      author: "Compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    });

    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
  });

  test("restores a field carrier after comparison inserts its result text", async () => {
    const base = await docxWith([{ type: "run", content: [{ type: "text", text: "Count: " }] }]);
    const target = await docxWith([
      { type: "run", content: [{ type: "text", text: "Count: " }] },
      {
        type: "simpleField",
        fieldType: "NUMWORDS",
        instruction: " NUMWORDS ",
        content: [{ type: "run", content: [{ type: "text", text: "12" }] }],
      },
    ]);

    const compared = await compareDocx(base, target, {
      author: "Compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    });

    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
  });

  test("removes a field carrier together with its result text", async () => {
    const base = await docxWith([
      { type: "run", content: [{ type: "text", text: "Count: " }] },
      {
        type: "simpleField",
        fieldType: "NUMWORDS",
        instruction: " NUMWORDS ",
        content: [{ type: "run", content: [{ type: "text", text: "12" }] }],
      },
    ]);
    const target = await docxWith([{ type: "run", content: [{ type: "text", text: "Count: " }] }]);

    const compared = await compareDocx(base, target, {
      author: "Compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    });

    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
  });

  test("tracks a field restored into replacement text through accept and reject", () => {
    const target = documentWith([schema.text("Pages: "), field()]);
    const state = EditorState.create({
      schema,
      doc: documentWith([schema.text("Pages: "), comparisonInsertedText("1")]),
    });
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
      schema.node("paragraph", null, [schema.text(" "), comparisonInsertedText("1")]),
    ]);
    const state = EditorState.create({ schema, doc: source });
    expect(resolve({ state, mode: "accept" }).doc.textContent).toBe("Pages: 1");
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
    expect(
      resolve({ state: reviewed, mode: "reject" }).doc.eq(resolve({ state, mode: "reject" }).doc),
    ).toBe(true);
  });

  test("preserves the target order for multiple atoms at one text boundary", () => {
    const target = documentWith([schema.text("Counts: "), field(), wordsField()]);
    const state = EditorState.create({
      schema,
      doc: documentWith([schema.text("Counts: "), comparisonInsertedText("12")]),
    });
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
    expect(resolve({ state: state.apply(result.transaction), mode: "accept" }).doc.eq(target)).toBe(
      true,
    );
  });

  test("tracks a zero-width page break without changing surrounding text", () => {
    const target = documentWith([
      schema.text("Before"),
      schema.node("pageBreakRun"),
      schema.text("After"),
    ]);
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

  test("ignores package-local comment ids beside an unchanged supported carrier", () => {
    const source = documentWith([
      schema.node("pageBreakRun"),
      schema.node("commentReference", { commentId: 3 }),
    ]);
    const target = documentWith([
      schema.node("pageBreakRun"),
      schema.node("commentReference", { commentId: 1 }),
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

    expect(result).toMatchObject({ status: "matched", rangeCount: 0 });
    expect(sameInlineAtoms(source, target)).toBe(true);
  });

  test("ignores unrelated preserved markup while restoring a supported carrier", () => {
    const source = documentWith([schema.text("AB")]);
    const target = documentWith([
      schema.text("A"),
      schema.node("pageBreakRun"),
      schema.text("B"),
      schema.node("preservedXml", {
        xml: '<w:proofErr w:type="gramStart"/>',
        text: "",
        level: "inline",
      }),
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
    const accepted = resolve({ state: state.apply(result.transaction), mode: "accept" }).doc;
    expect(sameInlineAtoms(accepted, target)).toBe(true);
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
      state: EditorState.create({
        schema,
        doc: documentWith([schema.text("Pages: "), comparisonInsertedText("1")]),
      }),
      targetSnapshot: createFolioAIEditSnapshot(target),
      revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
      originalRevisionIdSeed: 10,
      author: "Compare",
      maxRanges: 0,
    });

    expect(result.status).toBe("budget-exceeded");
  });
});
