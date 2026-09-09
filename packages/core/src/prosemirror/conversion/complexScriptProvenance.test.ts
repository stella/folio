import { describe, expect, test } from "bun:test";
import { EditorState, NodeSelection, TextSelection, type Command } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import type { Document, TextFormatting } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { setFontSize, toggleBold, toggleItalic } from "../commands/formatting";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const documentWithRun = (
  formatting: TextFormatting,
  inheritedFormatting?: TextFormatting,
): Document => {
  const document = createEmptyDocument();
  document.package.styles = {
    ...(inheritedFormatting ? { docDefaults: { rPr: inheritedFormatting } } : {}),
    styles: [],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      content: [
        {
          type: "run",
          formatting,
          content: [{ type: "text", text: "direct" }],
        },
      ],
    },
  ];
  return document;
};

const firstRunFormatting = (document: Document): TextFormatting | undefined => {
  const paragraph = document.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  return run?.type === "run" ? run.formatting : undefined;
};

const editableRoundTrip = (document: Document): TextFormatting | undefined =>
  firstRunFormatting(
    fromProseDoc(toProseDoc(document, { styles: document.package.styles }), document),
  );

const reopenThroughDocx = async (document: Document): Promise<Document> =>
  parseDocx(await createDocx(document), { detectVariables: false, preloadFonts: false });

const clonedEditorDocument = (document: Document) =>
  schema.nodeFromJSON(toProseDoc(document, { styles: document.package.styles }).toJSON());

const applyCommand = (state: EditorState, command: Command): EditorState => {
  let next = state;
  command(state, (transaction) => {
    next = state.apply(transaction);
  });
  return next;
};

const assertTwoPackageCycles = async (
  document: Document,
  expected: TextFormatting | undefined,
): Promise<void> => {
  const first = await reopenThroughDocx(document);
  expect(firstRunFormatting(first)).toEqual(expected);
  const second = await reopenThroughDocx(first);
  expect(firstRunFormatting(second)).toEqual(expected);
};

describe("complex-script run-property provenance", () => {
  test("preserves an absent szCs beside a direct size and an inherited complex size", () => {
    expect(editableRoundTrip(documentWithRun({ fontSize: 22 }, { fontSizeCs: 30 }))).toEqual({
      fontSize: 22,
    });
  });

  test("preserves an absent szCs beside a standalone direct size", () => {
    expect(editableRoundTrip(documentWithRun({ fontSize: 22 }))).toEqual({ fontSize: 22 });
  });

  test("does not materialize inherited ordinary or complex formatting on the run", () => {
    expect(
      editableRoundTrip(
        documentWithRun(
          {},
          {
            bold: true,
            boldCs: false,
            italic: true,
            italicCs: false,
            fontSize: 22,
            fontSizeCs: 30,
          },
        ),
      ),
    ).toBeUndefined();
  });

  test("preserves explicit complex properties equal to their ordinary partners", () => {
    expect(
      editableRoundTrip(
        documentWithRun({
          bold: true,
          boldCs: true,
          italic: true,
          italicCs: true,
          fontSize: 22,
          fontSizeCs: 22,
        }),
      ),
    ).toEqual({
      bold: true,
      boldCs: true,
      italic: true,
      italicCs: true,
      fontSize: 22,
      fontSizeCs: 22,
    });
  });

  test("a raw ordinary mark does not synthesize complex-script authorship", () => {
    const pmDocument = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("authored", [
          schema.mark("bold"),
          schema.mark("italic"),
          schema.mark("fontSize", { size: 22 }),
        ]),
      ]),
    ]);

    expect(firstRunFormatting(fromProseDoc(pmDocument))).toEqual({
      bold: true,
      italic: true,
      fontSize: 22,
    });
    expect(createFolioAIEditSnapshot(pmDocument).blocks.at(0)?.previewRuns).toEqual([
      {
        text: "authored",
        bold: true,
        italic: true,
        fontSizePt: 11,
        directFormatting: {
          bold: true,
          italic: true,
          fontSizePt: 11,
        },
      },
    ]);
  });

  test("toolbar commands explicitly author ordinary and complex-script companions", async () => {
    const base = documentWithRun({});
    let state = EditorState.create({ doc: clonedEditorDocument(base) });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 7)));
    state = applyCommand(state, toggleBold);
    state = applyCommand(state, toggleItalic);
    state = applyCommand(state, setFontSize(22));

    const authored = fromProseDoc(state.doc, base);
    const expected = {
      bold: true,
      boldCs: true,
      italic: true,
      italicCs: true,
      fontSize: 22,
      fontSizeCs: 22,
    };
    expect(firstRunFormatting(authored)).toEqual(expected);
    expect(createFolioAIEditSnapshot(state.doc).blocks.at(0)?.previewRuns).toEqual([
      {
        text: "direct",
        bold: true,
        italic: true,
        fontSizePt: 11,
        directFormatting: {
          bold: true,
          italic: true,
          fontSizePt: 11,
        },
      },
    ]);
    await assertTwoPackageCycles(authored, expected);
  });

  test.each([
    ["text", () => schema.text("x")],
    ["tab", () => schema.node("tab")],
    ["hard break", () => schema.node("hardBreak")],
    ["symbol", () => schema.node("symbol", { font: "Wingdings", char: "F06F" })],
    [
      "field",
      () =>
        schema.node("field", {
          fieldType: "PAGE",
          instruction: " PAGE ",
          displayText: "1",
          fieldKind: "simple",
          fldLock: false,
          dirty: false,
        }),
    ],
    [
      "structured field",
      () =>
        schema.node(
          "structuredField",
          {
            fieldType: "REF",
            instruction: " REF carrier ",
            displayText: "field",
            fieldKind: "simple",
            fldLock: false,
            dirty: false,
          },
          [schema.text("field")],
        ),
    ],
  ] as const)("toolbar commands preserve companion provenance on a selected %s", (_, makeNode) => {
    for (const command of [toggleBold, setFontSize(22)]) {
      const carrier = makeNode();
      const doc = schema.node("doc", null, [schema.node("paragraph", null, [carrier])]);
      let state = EditorState.create({ doc });
      state = state.apply(
        state.tr.setSelection(
          carrier.isText
            ? TextSelection.create(state.doc, 1, 2)
            : NodeSelection.create(state.doc, 1),
        ),
      );
      state = applyCommand(state, command);

      const formattedNodes: (typeof carrier)[] = [];
      state.doc.firstChild?.descendants((node) => {
        if (
          node.isInline &&
          (node.type.name !== "structuredField" || formattedNodes.length === 0)
        ) {
          formattedNodes.push(node);
        }
        return true;
      });
      for (const node of formattedNodes) {
        const override = node.marks.find(({ type }) => type.name === "runFormattingOverride");
        expect(override, `${node.type.name} must carry direct-formatting provenance`).toBeDefined();
        if (command === toggleBold) {
          expect(override?.attrs["bold"]).toBe(true);
          expect(override?.attrs["boldCs"]).toBe(true);
        } else {
          expect(override?.attrs["directFontProperties"]).toContain("fontSize");
          expect(override?.attrs["fontSizeCs"]).toBe(22);
        }
      }
    }
  });

  test("collapsed toolbar formatting keeps companion provenance in stored marks", () => {
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [schema.text("x")])]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
    state = applyCommand(state, toggleBold);

    const override = state.storedMarks?.find(({ type }) => type.name === "runFormattingOverride");
    expect(override?.attrs["bold"]).toBe(true);
    expect(override?.attrs["boldCs"]).toBe(true);
  });

  test("a nested field selection updates only the result text receiving visible formatting", () => {
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF carrier ",
        displayText: "AB",
        fieldKind: "simple",
        fldLock: false,
        dirty: false,
      },
      [schema.text("AB")],
    );
    let state = EditorState.create({
      doc: schema.node("doc", null, [schema.node("paragraph", null, [field])]),
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 3)));
    state = applyCommand(state, toggleBold);

    const updatedField = state.doc.firstChild?.firstChild;
    expect(updatedField?.marks.some(({ type }) => type.name === "bold")).toBe(false);
    expect(updatedField?.marks.some(({ type }) => type.name === "runFormattingOverride")).toBe(
      false,
    );
    expect(updatedField?.child(0).text).toBe("A");
    expect(updatedField?.child(0).marks.map(({ type }) => type.name)).toEqual([
      "bold",
      "runFormattingOverride",
    ]);
    expect(updatedField?.child(1).text).toBe("B");
    expect(updatedField?.child(1).marks).toEqual([]);
  });

  test.each([
    [
      "bookmark boundary",
      () => schema.node("bookmarkBoundary", { type: "start", id: 1, name: "carrier" }),
    ],
    ["image", () => schema.node("image", { src: "data:image/png;base64," })],
    ["math", () => schema.node("math")],
    ["rendered page break", () => schema.node("renderedPageBreak")],
    ["shape", () => schema.node("shape")],
    ["text-box anchor", () => schema.node("textBoxAnchor", { anchorId: "carrier" })],
  ] as const)("toolbar formatting leaves %s without run provenance", (_, makeNode) => {
    const atom = makeNode();
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [atom])]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, 1)));
    state = applyCommand(state, toggleBold);

    expect(
      state.doc.firstChild?.firstChild?.marks.some(
        ({ type }) => type.name === "runFormattingOverride",
      ),
    ).toBe(false);
  });

  test.each([
    ["ordinary bold", { bold: true }],
    ["single underline", { underline: { style: "single" } }],
  ] as const)(
    "keeps inferable %s carrierless through editor and package cycles",
    async (_, formatting) => {
      const document = documentWithRun(formatting);
      const proseDocument = clonedEditorDocument(document);
      const json = JSON.stringify(proseDocument.toJSON());

      expect(json).not.toContain("_authoredOn");
      expect(json).not.toContain("_authoredOff");
      expect(json).not.toContain("_authoredValues");

      const saved = fromProseDoc(proseDocument, document);
      expect(firstRunFormatting(saved)).toEqual(formatting);
      await assertTwoPackageCycles(saved, formatting);
    },
  );

  test.each([
    ["same-valued inherited bold", { bold: true }, { bold: true }],
    ["underline none", { underline: { style: "none" } }, undefined],
    [
      "partial font family",
      { fontFamily: { ascii: "Direct" } },
      { fontFamily: { ascii: "Inherited", hAnsi: "Inherited" } },
    ],
    ["theme color", { color: { themeColor: "accent1" } }, undefined],
  ] as const)(
    "keeps ambiguous %s in the sparse authored carrier",
    async (_, formatting, inheritedFormatting) => {
      const document = documentWithRun(formatting, inheritedFormatting);
      const proseDocument = clonedEditorDocument(document);
      const json = JSON.stringify(proseDocument.toJSON());

      expect(json).toContain("_authored");

      const saved = fromProseDoc(proseDocument, document);
      expect(firstRunFormatting(saved)).toEqual(formatting);
      await assertTwoPackageCycles(saved, formatting);
    },
  );

  test("raw mark transactions update a carrierless direct property before save", async () => {
    const document = documentWithRun({ bold: true });
    const state = EditorState.create({ doc: clonedEditorDocument(document) });
    const bold = schema.marks["bold"];
    if (!bold) {
      throw new Error("Expected bold mark type");
    }
    const withoutBold = state.apply(state.tr.removeMark(1, 7, bold)).doc;
    const cleared = fromProseDoc(withoutBold, document);

    expect(firstRunFormatting(cleared)).toBeUndefined();
    await assertTwoPackageCycles(cleared, undefined);

    const plain = documentWithRun({});
    const plainState = EditorState.create({ doc: clonedEditorDocument(plain) });
    const withBold = plainState.apply(plainState.tr.addMark(1, 7, bold.create())).doc;
    const authored = fromProseDoc(withBold, plain);

    expect(firstRunFormatting(authored)).toEqual({ bold: true });
    await assertTwoPackageCycles(authored, { bold: true });
  });
});
