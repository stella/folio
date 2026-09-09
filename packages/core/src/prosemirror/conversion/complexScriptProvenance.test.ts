import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection, type Command } from "prosemirror-state";

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
