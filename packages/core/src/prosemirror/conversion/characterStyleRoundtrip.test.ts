/**
 * Character style (w:rStyle) round-trip through ProseMirror.
 *
 * A run's character style reference must survive load → edit → save: the
 * style's formatting is resolved through the style chain for rendering
 * (flattened into regular marks), while a compact `characterStyle` mark carries
 * the reference so the serializer re-emits `w:rStyle`.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { parseDocumentBody } from "../../docx/documentParser";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { serializeDocument } from "../../docx/serializer/documentSerializer";
import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import type { Document, Paragraph, Run, StyleDefinitions } from "../../types/document";
import { schema } from "../schema";
import { acceptAllChanges, rejectAllChanges } from "../commands/comments";
import { applyFormatMarks, captureFormatMarks } from "../commands/formatPainter";
import { createDocumentStylesPlugin } from "../plugins/documentStyles";
import { fromProseDoc, marksToTextFormatting } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const runText = (text: string, formatting?: Run["formatting"]): Run => {
  const run: Run = {
    type: "run",
    content: [{ type: "text", text }],
  };
  if (formatting) {
    run.formatting = formatting;
  }
  return run;
};

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: {
    document: {
      content: [paragraph],
    },
  },
});

const wrap = (...runs: Run[]): Document => wrapParagraph({ type: "paragraph", content: runs });

const styles: StyleDefinitions = {
  styles: [
    {
      styleId: "DefinedTerm",
      type: "character",
      name: "Defined Term",
      rPr: { italic: true, color: { rgb: "336699" } },
    },
    {
      styleId: "AccentChar",
      type: "character",
      name: "Accent Character",
      rPr: { color: { rgb: "00AA00" } },
    },
    {
      styleId: "StrongHeading",
      type: "paragraph",
      name: "Strong Heading",
      rPr: { bold: true, color: { rgb: "0000FF" } },
    },
    {
      styleId: "ToggleHeading",
      type: "paragraph",
      name: "Toggle Heading",
      rPr: { bold: true },
    },
    {
      styleId: "PlainParagraph",
      type: "paragraph",
      name: "Plain Paragraph",
    },
    {
      styleId: "StrongCharacter",
      type: "character",
      name: "Strong Character",
      rPr: { bold: true },
    },
    {
      styleId: "LatinEmphasis",
      type: "character",
      name: "Latin Emphasis",
      rPr: { bold: true, boldCs: false, italic: true, italicCs: false },
    },
    {
      styleId: "DefaultStrongCharacter",
      type: "character",
      name: "Default Strong Character",
      rPr: { bold: true },
    },
  ],
};

const stylesWithDefaultCharacter: StyleDefinitions = {
  styles: styles.styles.map((style) =>
    style.styleId === "DefaultStrongCharacter" ? { ...style, default: true } : style,
  ),
};

const withStyles = (document: Document, styleDefinitions: StyleDefinitions = styles): Document => {
  document.package.styles = styleDefinitions;
  return document;
};

const reopenThroughDocx = async (document: Document): Promise<Document> =>
  parseDocx(await createDocx(document), { detectVariables: false, preloadFonts: false });

const firstParagraph = (document: Document): Paragraph => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("Expected first block to be a paragraph");
  }
  return block;
};

const paragraphRuns = (paragraph: Paragraph): Run[] =>
  paragraph.content.filter((content): content is Run => content.type === "run");

const findRun = (paragraph: Paragraph, text: string): Run => {
  for (const run of paragraphRuns(paragraph)) {
    const concat = run.content
      .filter((c) => c.type === "text")
      .map((c) => c.text)
      .join("");
    if (concat === text) {
      return run;
    }
  }
  throw new Error(`Expected run with text ${text}`);
};

const findTextNode = (doc: PMNode, text: string) => {
  let found:
    | {
        from: number;
        node: PMNode;
        to: number;
      }
    | undefined;
  doc.descendants((node, pos) => {
    if (!found && node.isText && node.text === text) {
      found = { from: pos, node, to: pos + node.nodeSize };
    }
    return found === undefined;
  });
  if (!found) {
    throw new Error(`Expected PM text node ${text}`);
  }
  return found;
};

const paragraphAt = (document: Document, index: number): Paragraph => {
  const block = document.package.document.content.at(index);
  if (block?.type !== "paragraph") {
    throw new Error(`Expected paragraph ${index}`);
  }
  return block;
};

const selectText = (state: EditorState, text: string): EditorState => {
  let range: { from: number; to: number } | null = null;
  state.doc.descendants((node, position) => {
    if (!range && node.isText && node.text === text) {
      range = { from: position, to: position + node.nodeSize };
    }
  });
  if (!range) {
    throw new Error(`Expected PM text node ${text}`);
  }
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, range.from, range.to)));
};

const reopenSerializedBody = (document: Document): Document => ({
  package: {
    ...document.package,
    document: parseDocumentBody(serializeDocument(document)),
  },
});

const createStyledEditorState = (
  document: Document,
  styleDefinitions: StyleDefinitions,
): EditorState =>
  EditorState.create({
    doc: toProseDoc(document, { styles: styleDefinitions }),
    plugins: [createDocumentStylesPlugin(styleDefinitions)],
  });
const markNames = (document: Document, text: string): string[] => {
  const pmDoc = toProseDoc(document, { styles });
  let names: string[] | undefined;
  pmDoc.descendants((node) => {
    if (node.isText && node.text === text) {
      names = node.marks.map((mark) => mark.type.name);
    }
    return true;
  });
  if (!names) {
    throw new Error(`Expected PM text node ${text}`);
  }
  return names;
};

describe("characterStyle mark schema registration", () => {
  test("schema includes the characterStyle mark", () => {
    expect(schema.marks["characterStyle"]).toBeDefined();
  });

  test("run override schema exposes no raw source-formatting payloads", () => {
    const attrs = Object.keys(schema.marks["runFormattingOverride"]?.spec.attrs ?? {});
    expect(attrs).not.toContain("_baseRPr");
    expect(attrs).not.toContain("_effectiveRPr");
    expect(attrs).not.toContain("_directRPr");
    expect(attrs).not.toContain("_sourceStyleId");
  });
});

describe("resolver-less snapshots", () => {
  test("keep effective character-style visuals without presenting them as direct formatting", () => {
    const styleDefinitions: StyleDefinitions = {
      styles: [
        {
          styleId: "SnapshotCharacter",
          type: "character",
          rPr: {
            bold: true,
            boldCs: true,
            fontSize: 30,
            fontSizeCs: 30,
            color: { rgb: "FF0000" },
          },
        },
      ],
    };
    const document = withStyles(
      wrap(runText("Styled", { styleId: "SnapshotCharacter" })),
      styleDefinitions,
    );

    const block = createFolioAIEditSnapshot(
      toProseDoc(document, { styles: styleDefinitions }),
    ).blocks.at(0);

    expect(block?.previewRuns).toEqual([
      {
        text: "Styled",
        effectiveFormatting: {
          bold: true,
          color: { rgb: "FF0000" },
          fontSize: 30,
          styleId: "SnapshotCharacter",
        },
        authoredFormatting: { styleId: "SnapshotCharacter" },
      },
    ]);
  });

  test.each([
    {
      label: "ordinary visual formatting",
      marks: [schema.mark("characterStyle", { styleId: "Character" }), schema.mark("bold")],
      expected: { bold: true, styleId: "Character" },
    },
    {
      label: "an independent complex-script override",
      marks: [
        schema.mark("characterStyle", { styleId: "Character" }),
        schema.mark("bold"),
        schema.mark("runFormattingOverride", { boldCs: false }),
      ],
      expected: { bold: true, boldCs: false, styleId: "Character" },
    },
    {
      label: "explicit ordinary off and complex-script on",
      marks: [
        schema.mark("characterStyle", { styleId: "Character" }),
        schema.mark("runFormattingOverride", { bold: false, boldCs: true }),
      ],
      expected: { bold: false, boldCs: true, styleId: "Character" },
    },
  ])(
    "conservatively saves $label on a character-styled run without a resolver",
    ({ marks, expected }) => {
      const pmDoc = schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("Styled", marks)]),
      ]);

      expect(findRun(firstParagraph(fromProseDoc(pmDoc)), "Styled").formatting).toEqual(expected);
    },
  );
});

describe("character style rendering resolution", () => {
  test("styled run renders the style's formatting via regular marks", () => {
    const input = wrap(runText("Term", { styleId: "DefinedTerm" }));
    const names = markNames(input, "Term");
    expect(names).toContain("italic");
    expect(names).toContain("textColor");
    expect(names).toContain("characterStyle");
  });

  test("direct formatting wins over the character style", () => {
    const input = wrap(runText("Term", { styleId: "DefinedTerm", color: { rgb: "FF0000" } }));
    const pmDoc = toProseDoc(input, { styles });
    let rgb: unknown;
    pmDoc.descendants((node) => {
      if (node.isText && node.text === "Term") {
        const mark = node.marks.find((m) => m.type.name === "textColor");
        rgb = mark?.attrs["rgb"];
      }
      return true;
    });
    expect(rgb).toBe("FF0000");
  });

  test("character style wins over the paragraph style", () => {
    const input = wrapParagraph({
      type: "paragraph",
      formatting: { styleId: "StrongHeading" },
      content: [runText("Term", { styleId: "AccentChar" })],
    });
    const pmDoc = toProseDoc(input, { styles });
    let rgb: unknown;
    let bold = false;
    pmDoc.descendants((node) => {
      if (node.isText && node.text === "Term") {
        const mark = node.marks.find((m) => m.type.name === "textColor");
        rgb = mark?.attrs["rgb"];
        bold = node.marks.some((m) => m.type.name === "bold");
      }
      return true;
    });
    // Color comes from the character style; bold still cascades down from
    // the paragraph style because the character style does not redefine it.
    expect(rgb).toBe("00AA00");
    expect(bold).toBe(true);
  });
});

describe("character style round-trip", () => {
  test("a pure style reference round-trips without baked direct formatting", () => {
    const input = withStyles(wrap(runText("Term", { styleId: "DefinedTerm" })));
    const pmDoc = toProseDoc(input, { styles });
    const out = fromProseDoc(pmDoc, input);
    const run = findRun(firstParagraph(out), "Term");
    expect(run.formatting).toEqual({ styleId: "DefinedTerm" });
  });

  test("direct overrides survive next to the style reference", () => {
    const input = withStyles(
      wrap(runText("Term", { styleId: "DefinedTerm", color: { rgb: "FF0000" } })),
    );
    const pmDoc = toProseDoc(input, { styles });
    const out = fromProseDoc(pmDoc, input);
    const run = findRun(firstParagraph(out), "Term");
    expect(run.formatting?.styleId).toBe("DefinedTerm");
    expect(run.formatting?.color).toEqual({ rgb: "FF0000" });
    // Italic came purely from the style — it must not be baked in.
    expect(run.formatting?.italic).toBeUndefined();
  });

  test("preserves an authored underline equal to its character style across DOCX reopen", async () => {
    const underline = { style: "double", color: { rgb: "FF0000" } } as const;
    const styleDefinitions: StyleDefinitions = {
      styles: [
        {
          styleId: "DoubleUnderline",
          type: "character",
          rPr: { underline },
        },
      ],
    };
    const input = withStyles(
      wrap(runText("Term", { styleId: "DoubleUnderline", underline })),
      styleDefinitions,
    );

    const once = await reopenThroughDocx(
      fromProseDoc(toProseDoc(input, { styles: styleDefinitions }), input),
    );
    const twice = await reopenThroughDocx(
      fromProseDoc(toProseDoc(once, { styles: once.package.styles }), once),
    );

    expect(findRun(firstParagraph(once), "Term").formatting).toEqual({
      styleId: "DoubleUnderline",
      underline,
    });
    expect(findRun(firstParagraph(twice), "Term").formatting).toEqual({
      styleId: "DoubleUnderline",
      underline,
    });
  });

  test("round-trip is stable across a second load/save cycle", () => {
    const input = withStyles(wrap(runText("Term", { styleId: "DefinedTerm" })));
    const once = fromProseDoc(toProseDoc(input, { styles }), input);
    const twice = fromProseDoc(toProseDoc(once, { styles }), once);
    expect(findRun(firstParagraph(twice), "Term").formatting).toEqual({
      styleId: "DefinedTerm",
    });
  });

  test("matching paragraph and explicit character toggles preserve only the style reference", () => {
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText("Term", { styleId: "StrongCharacter" })],
      }),
    );

    const proseDoc = toProseDoc(input, { styles });
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
    const out = fromProseDoc(clonedProseDoc, input);

    expect(findRun(firstParagraph(out), "Term").formatting).toEqual({
      styleId: "StrongCharacter",
    });
  });

  test("independent complex-script style toggles survive a JSON clone without direct formatting", () => {
    const input = withStyles(wrap(runText("Term", { styleId: "LatinEmphasis" })));

    const proseDoc = toProseDoc(input, { styles });
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
    const out = fromProseDoc(clonedProseDoc, input);

    expect(findRun(firstParagraph(out), "Term").formatting).toEqual({
      styleId: "LatinEmphasis",
    });
  });

  test("all inherited, character, and direct ordinary/CS toggle combinations round-trip", () => {
    const values = [false, true] as const;
    const directValues = [undefined, false, true] as const;
    for (const inheritedOrdinary of values) {
      for (const inheritedCs of values) {
        for (const characterOrdinary of values) {
          for (const characterCs of values) {
            for (const direct of directValues) {
              const matrixStyles: StyleDefinitions = {
                styles: [
                  {
                    styleId: "P",
                    type: "paragraph",
                    name: "Paragraph",
                    rPr: {
                      bold: inheritedOrdinary,
                      boldCs: inheritedCs,
                      italic: inheritedOrdinary,
                      italicCs: inheritedCs,
                    },
                  },
                  {
                    styleId: "C",
                    type: "character",
                    name: "Character",
                    rPr: {
                      bold: characterOrdinary,
                      boldCs: characterCs,
                      italic: characterOrdinary,
                      italicCs: characterCs,
                    },
                  },
                ],
              };
              const directFormatting =
                direct === undefined
                  ? { styleId: "C" }
                  : {
                      styleId: "C",
                      bold: direct,
                      boldCs: direct,
                      italic: direct,
                      italicCs: direct,
                    };
              const input = withStyles(
                wrapParagraph({
                  type: "paragraph",
                  formatting: { styleId: "P" },
                  content: [runText("Term", directFormatting)],
                }),
                matrixStyles,
              );

              const proseDoc = toProseDoc(input, { styles: matrixStyles });
              const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
              const out = fromProseDoc(clonedProseDoc, input);
              const matrixCase = [
                inheritedOrdinary,
                inheritedCs,
                characterOrdinary,
                characterCs,
                direct,
              ].join("/");

              expect({
                formatting: findRun(firstParagraph(out), "Term").formatting,
                matrixCase,
              }).toEqual({ formatting: directFormatting, matrixCase });
            }
          }
        }
      }
    }
  });

  test("authored direct positives remain fixed after a paragraph-style change", () => {
    const initialStyles: StyleDefinitions = {
      styles: [
        {
          styleId: "P",
          type: "paragraph",
          name: "Paragraph",
          rPr: { bold: true, boldCs: true, italic: true, italicCs: true },
        },
        {
          styleId: "C",
          type: "character",
          name: "Character",
          rPr: { bold: false, boldCs: false, italic: false, italicCs: false },
        },
      ],
    };
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "P" },
        content: [
          runText("Term", {
            styleId: "C",
            bold: true,
            boldCs: true,
            italic: true,
            italicCs: true,
          }),
        ],
      }),
      initialStyles,
    );
    const clonedProseDoc = schema.nodeFromJSON(
      toProseDoc(input, { styles: initialStyles }).toJSON(),
    );
    const saved = fromProseDoc(clonedProseDoc, input);
    const changedStyles: StyleDefinitions = {
      styles: [
        {
          styleId: "P",
          type: "paragraph",
          name: "Paragraph",
          rPr: { bold: false, boldCs: false, italic: false, italicCs: false },
        },
        {
          styleId: "C",
          type: "character",
          name: "Character",
          rPr: { bold: false, boldCs: false, italic: false, italicCs: false },
        },
      ],
    };
    const changedProseDoc = schema.nodeFromJSON(
      toProseDoc(saved, { styles: changedStyles }).toJSON(),
    );
    const changedFlowParagraph = toFlowBlocks(changedProseDoc, {}).find(
      (block) => block.kind === "paragraph",
    );
    const changedFlowRun = changedFlowParagraph?.runs.find((run) => run.kind === "text");

    expect(findRun(firstParagraph(saved), "Term").formatting).toEqual({
      styleId: "C",
      bold: true,
      boldCs: true,
      italic: true,
      italicCs: true,
    });
    expect(changedFlowRun).toMatchObject({
      bold: true,
      complexScriptBold: true,
      italic: true,
      complexScriptItalic: true,
    });
  });

  test("an authored direct off survives when the style cascade is already off", () => {
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText("Term", { styleId: "StrongCharacter", bold: false })],
      }),
    );

    const proseDoc = toProseDoc(input, { styles });
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
    const out = fromProseDoc(clonedProseDoc, input);

    expect(findRun(firstParagraph(out), "Term").formatting).toEqual({
      styleId: "StrongCharacter",
      bold: false,
    });
  });

  test("matching paragraph and default character toggles add no direct formatting", () => {
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText("Term")],
      }),
      stylesWithDefaultCharacter,
    );

    const proseDoc = toProseDoc(input, { styles: stylesWithDefaultCharacter });
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
    const out = fromProseDoc(clonedProseDoc, input);

    expect(findRun(firstParagraph(out), "Term").formatting).toBeUndefined();
  });

  test("a later paragraph-style change is not masked by a synthesized direct off", () => {
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText("Term", { styleId: "StrongCharacter" })],
      }),
    );
    const saved = fromProseDoc(toProseDoc(input, { styles }), input);
    const changedStyles: StyleDefinitions = {
      styles: styles.styles.map((style) =>
        style.styleId === "ToggleHeading" ? { ...style, rPr: { bold: false } } : style,
      ),
    };

    const names = toProseDoc(saved, { styles: changedStyles })
      .child(0)
      .child(0)
      .marks.map((mark) => mark.type.name);

    expect(findRun(firstParagraph(saved), "Term").formatting).toEqual({
      styleId: "StrongCharacter",
    });
    expect(names).toContain("bold");
  });

  test("a later paragraph-style change still reaches the implicit default character style", () => {
    const input = withStyles(
      wrapParagraph({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText("Term")],
      }),
      stylesWithDefaultCharacter,
    );
    const saved = fromProseDoc(toProseDoc(input, { styles: stylesWithDefaultCharacter }), input);
    const changedStyles: StyleDefinitions = {
      styles: stylesWithDefaultCharacter.styles.map((style) =>
        style.styleId === "ToggleHeading" ? { ...style, rPr: { bold: false } } : style,
      ),
    };

    const names = toProseDoc(saved, { styles: changedStyles })
      .child(0)
      .child(0)
      .marks.map((mark) => mark.type.name);

    expect(findRun(firstParagraph(saved), "Term").formatting).toBeUndefined();
    expect(names).toContain("bold");
  });

  test("serialized editor state stays linear for many inherited toggle runs", () => {
    const paragraphs = Array.from(
      { length: 1000 },
      (_, index): Paragraph => ({
        type: "paragraph",
        formatting: { styleId: "ToggleHeading" },
        content: [runText(`Term ${index}`)],
      }),
    );
    const plain: Document = {
      package: {
        document: { content: paragraphs.map(({ content, type }) => ({ content, type })) },
      },
    };
    const styled: Document = {
      package: { document: { content: paragraphs }, styles: stylesWithDefaultCharacter },
    };

    const plainSize = JSON.stringify(toProseDoc(plain).toJSON()).length;
    const styledProseDoc = toProseDoc(styled, { styles: stylesWithDefaultCharacter });
    const styledJson = JSON.stringify(styledProseDoc.toJSON());
    const saved = fromProseDoc(styledProseDoc, styled);

    expect(styledJson.length).toBeLessThan(plainSize * 2);
    expect(styledJson).not.toContain("_effectiveRPr");
    expect(styledJson).not.toContain("_directRPr");
    expect(styledJson).not.toContain("runFormattingOverride");
    expect(saved.package.document.content).toHaveLength(1000);
    expect(findRun(firstParagraph(saved), "Term 0").formatting).toBeUndefined();
  });

  test("painting an effective off across paragraph styles saves a direct off", async () => {
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "ToggleHeading" },
              content: [runText("From", { styleId: "StrongCharacter" })],
            },
            {
              type: "paragraph",
              formatting: { styleId: "PlainParagraph" },
              content: [runText("To")],
            },
          ],
        },
        styles,
      },
    };
    let state = createStyledEditorState(input, styles);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 5)));
    const captured = captureFormatMarks(state);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7, 9)));
    applyFormatMarks(captured)(state, (transaction) => {
      state = state.apply(transaction);
    });

    const saved = fromProseDoc(state.doc, input);
    const target = saved.package.document.content.at(1);
    if (target?.type !== "paragraph") {
      throw new Error("Expected target paragraph");
    }

    expect(findRun(target, "To").formatting).toEqual({
      styleId: "StrongCharacter",
      bold: false,
    });

    const reopened = await reopenThroughDocx(saved);
    const reopenedTarget = reopened.package.document.content.at(1);
    if (reopenedTarget?.type !== "paragraph") {
      throw new Error("Expected reopened target paragraph");
    }
    expect(findRun(reopenedTarget, "To").formatting).toEqual({
      styleId: "StrongCharacter",
      bold: false,
    });
  });

  test("painting an effective on recomputes the character style in the target paragraph", async () => {
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "PlainParagraph" },
              content: [runText("From", { styleId: "StrongCharacter" })],
            },
            {
              type: "paragraph",
              formatting: { styleId: "ToggleHeading" },
              content: [runText("To")],
            },
          ],
        },
        styles,
      },
    };
    let state = createStyledEditorState(input, styles);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 5)));
    const captured = captureFormatMarks(state);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7, 9)));
    applyFormatMarks(captured)(state, (transaction) => {
      state = state.apply(transaction);
    });

    const saved = fromProseDoc(state.doc, input);
    const target = saved.package.document.content.at(1);
    if (target?.type !== "paragraph") {
      throw new Error("Expected target paragraph");
    }
    expect(findRun(target, "To").formatting).toEqual({
      styleId: "StrongCharacter",
      bold: true,
    });

    const reopened = await reopenThroughDocx(saved);
    const reopenedTarget = reopened.package.document.content.at(1);
    if (reopenedTarget?.type !== "paragraph") {
      throw new Error("Expected reopened target paragraph");
    }
    expect(findRun(reopenedTarget, "To").formatting).toEqual({
      styleId: "StrongCharacter",
      bold: true,
    });
  });

  test.each([
    {
      label: "plain into bold",
      sourceStyleId: "PlainParagraph",
      targetStyleId: "ToggleHeading",
      expectedFormatting: { bold: false },
    },
    {
      label: "bold into plain",
      sourceStyleId: "ToggleHeading",
      targetStyleId: "PlainParagraph",
      expectedFormatting: { bold: true },
    },
  ] as const)(
    "copying a visually $label paragraph preserves appearance after save and reopen",
    async ({ sourceStyleId, targetStyleId, expectedFormatting }) => {
      const input: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                formatting: { styleId: sourceStyleId },
                content: [runText("Source")],
              },
              {
                type: "paragraph",
                formatting: { styleId: targetStyleId },
                content: [runText("Target")],
              },
            ],
          },
          styles,
        },
      };
      const state = EditorState.create({ doc: toProseDoc(input, { styles }) });
      const source = findTextNode(state.doc, "Source");
      const target = findTextNode(state.doc, "Target");
      const copied = state.apply(
        state.tr.replaceWith(
          target.from,
          target.to,
          schema.text(source.node.text ?? "", source.node.marks),
        ),
      );

      const saved = fromProseDoc(copied.doc, input);
      const targetParagraph = saved.package.document.content.at(1);
      if (targetParagraph?.type !== "paragraph") {
        throw new Error("Expected target paragraph");
      }
      expect(findRun(targetParagraph, "Source").formatting).toEqual(expectedFormatting);

      const reopened = await reopenThroughDocx(saved);
      const reopenedTarget = reopened.package.document.content.at(1);
      if (reopenedTarget?.type !== "paragraph") {
        throw new Error("Expected reopened target paragraph");
      }
      expect(findRun(reopenedTarget, "Source").formatting).toEqual(expectedFormatting);
    },
  );

  test("removing the only inherited toggle mark saves an explicit off", async () => {
    const input = wrapParagraph({
      type: "paragraph",
      formatting: { styleId: "ToggleHeading" },
      content: [runText("Term")],
    });
    const state = EditorState.create({ doc: toProseDoc(input, { styles }) });
    const bold = schema.marks["bold"];
    if (!bold) {
      throw new Error("Expected bold mark type");
    }
    const term = findTextNode(state.doc, "Term");
    const edited = state.apply(state.tr.removeMark(term.from, term.to, bold));

    const saved = fromProseDoc(edited.doc, input);
    expect(findRun(firstParagraph(saved), "Term").formatting).toEqual({ bold: false });

    const reopened = await reopenThroughDocx(saved);
    expect(findRun(firstParagraph(reopened), "Term").formatting).toEqual({ bold: false });
  });

  test("painting independent complex-script style toggles emits no redundant direct offs", () => {
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "PlainParagraph" },
              content: [runText("From", { styleId: "LatinEmphasis" })],
            },
            {
              type: "paragraph",
              formatting: { styleId: "PlainParagraph" },
              content: [runText("To")],
            },
          ],
        },
        styles,
      },
    };
    let state = createStyledEditorState(input, styles);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 5)));
    const captured = captureFormatMarks(state);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7, 9)));
    applyFormatMarks(captured)(state, (transaction) => {
      state = state.apply(transaction);
    });

    const saved = fromProseDoc(state.doc, input);
    const target = saved.package.document.content.at(1);
    if (target?.type !== "paragraph") {
      throw new Error("Expected target paragraph");
    }

    expect(findRun(target, "To").formatting).toEqual({ styleId: "LatinEmphasis" });
  });

  test("preserves effective ordinary and complex-script toggles across every style context", () => {
    const togglePairs = [
      { ordinary: false, complex: false },
      { ordinary: false, complex: true },
      { ordinary: true, complex: false },
      { ordinary: true, complex: true },
    ] as const;
    for (const sourceParagraph of togglePairs) {
      for (const sourceCharacter of togglePairs) {
        for (const targetParagraph of togglePairs) {
          for (const targetCharacter of togglePairs) {
            const matrixStyles: StyleDefinitions = {
              docDefaults: {
                rPr: {
                  bold: false,
                  boldCs: false,
                  fontFamily: { asciiTheme: "minorHAnsi", hAnsiTheme: "minorHAnsi" },
                  fontSize: 22,
                  fontSizeCs: 22,
                  italic: false,
                  italicCs: false,
                },
              },
              styles: [
                {
                  styleId: "SourceParagraph",
                  type: "paragraph",
                  name: "Source Paragraph",
                  rPr: {
                    bold: sourceParagraph.ordinary,
                    boldCs: sourceParagraph.complex,
                    italic: sourceParagraph.ordinary,
                    italicCs: sourceParagraph.complex,
                  },
                },
                {
                  styleId: "SourceCharacter",
                  type: "character",
                  name: "Source Character",
                  rPr: {
                    bold: sourceCharacter.ordinary,
                    boldCs: sourceCharacter.complex,
                    italic: sourceCharacter.ordinary,
                    italicCs: sourceCharacter.complex,
                  },
                },
                {
                  styleId: "TargetParagraph",
                  type: "paragraph",
                  name: "Target Paragraph",
                  rPr: {
                    bold: targetParagraph.ordinary,
                    boldCs: targetParagraph.complex,
                    italic: targetParagraph.ordinary,
                    italicCs: targetParagraph.complex,
                  },
                },
                {
                  styleId: "TargetCharacter",
                  type: "character",
                  name: "Target Character",
                  rPr: {
                    bold: targetCharacter.ordinary,
                    boldCs: targetCharacter.complex,
                    italic: targetCharacter.ordinary,
                    italicCs: targetCharacter.complex,
                  },
                },
              ],
            };
            const input: Document = {
              package: {
                document: {
                  content: [
                    {
                      type: "paragraph",
                      formatting: { styleId: "SourceParagraph" },
                      content: [runText("From", { styleId: "SourceCharacter" })],
                    },
                    {
                      type: "paragraph",
                      formatting: { styleId: "TargetParagraph" },
                      content: [runText("To", { styleId: "TargetCharacter" })],
                    },
                  ],
                },
                styles: matrixStyles,
              },
            };
            let state = createStyledEditorState(input, matrixStyles);
            state = selectText(state, "From");
            const source = captureFormatMarks(state);
            state = selectText(state, "To");
            applyFormatMarks(source)(state, (transaction) => {
              state = state.apply(transaction);
            });

            const saved = fromProseDoc(state.doc, input);
            const reopened = reopenSerializedBody(saved);
            let reopenedState = createStyledEditorState(reopened, matrixStyles);
            reopenedState = selectText(reopenedState, "To");
            const target = captureFormatMarks(reopenedState);
            const matrixCase = [sourceParagraph, sourceCharacter, targetParagraph, targetCharacter]
              .map(({ ordinary, complex }) => `${Number(ordinary)}${Number(complex)}`)
              .join("/");

            expect({
              actual: {
                bold: target.effectiveFormatting.bold ?? false,
                boldCs: target.effectiveFormatting.boldCs ?? false,
                italic: target.effectiveFormatting.italic ?? false,
                italicCs: target.effectiveFormatting.italicCs ?? false,
              },
              matrixCase,
            }).toEqual({
              actual: {
                bold: source.effectiveFormatting.bold ?? false,
                boldCs: source.effectiveFormatting.boldCs ?? false,
                italic: source.effectiveFormatting.italic ?? false,
                italicCs: source.effectiveFormatting.italicCs ?? false,
              },
              matrixCase,
            });
            expect(findRun(paragraphAt(saved, 1), "To").formatting?.styleId).toBe(
              "SourceCharacter",
            );
          }
        }
      }
    }
  });

  test("doc-default-active character-style toggles stay active in layout after save and reopen", async () => {
    const defaultsActiveStyles: StyleDefinitions = {
      docDefaults: {
        rPr: { bold: true, boldCs: true, italic: true, italicCs: true },
      },
      styles: [
        {
          styleId: "DefaultActiveParagraph",
          type: "paragraph",
          name: "Default Active Paragraph",
          rPr: { bold: true, boldCs: true, italic: true, italicCs: true },
        },
        {
          styleId: "DefaultActiveCharacter",
          type: "character",
          name: "Default Active Character",
          rPr: { bold: true, boldCs: true, italic: true, italicCs: true },
        },
      ],
    };
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "DefaultActiveParagraph" },
              content: [runText("Defaults stay active", { styleId: "DefaultActiveCharacter" })],
            },
          ],
        },
        styles: defaultsActiveStyles,
      },
    };
    const proseDoc = toProseDoc(input, { styles: defaultsActiveStyles });
    const beforeSave = toFlowBlocks(proseDoc, { styles: defaultsActiveStyles })
      .flatMap((block) => (block.kind === "paragraph" ? block.runs : []))
      .find((run) => run.kind === "text");

    expect(beforeSave).toMatchObject({
      bold: true,
      complexScriptBold: true,
      complexScriptItalic: true,
      italic: true,
    });

    const saved = fromProseDoc(proseDoc, input);
    expect(findRun(firstParagraph(saved), "Defaults stay active").formatting).toEqual({
      styleId: "DefaultActiveCharacter",
    });

    const reopened = await reopenThroughDocx(saved);
    expect(findRun(firstParagraph(reopened), "Defaults stay active").formatting).toEqual({
      styleId: "DefaultActiveCharacter",
    });
    const afterReopen = toFlowBlocks(toProseDoc(reopened, { styles: defaultsActiveStyles }), {
      styles: defaultsActiveStyles,
    })
      .flatMap((block) => (block.kind === "paragraph" ? block.runs : []))
      .find((run) => run.kind === "text");
    expect(afterReopen).toMatchObject({
      bold: true,
      complexScriptBold: true,
      complexScriptItalic: true,
      italic: true,
    });
  });

  test("keeps the complete safe target-relative cancellation set after serialize and reopen", () => {
    const plainFormatting = {
      allCaps: false,
      bold: false,
      boldCs: false,
      color: { auto: true },
      cs: false,
      doubleStrike: false,
      effect: "none" as const,
      emboss: false,
      emphasisMark: "none" as const,
      fontFamily: { asciiTheme: "minorHAnsi" as const, hAnsiTheme: "minorHAnsi" },
      fontSize: 22,
      fontSizeCs: 22,
      highlight: "none" as const,
      imprint: false,
      italic: false,
      italicCs: false,
      kerning: 0,
      outline: false,
      position: 0,
      scale: 100,
      shading: { pattern: "nil" as const },
      shadow: false,
      smallCaps: false,
      spacing: 0,
      strike: false,
      underline: { style: "none" as const },
      vertAlign: "baseline" as const,
    };
    const contextualFormatting = {
      allCaps: true,
      bold: true,
      boldCs: true,
      color: { rgb: "336699" },
      cs: true,
      doubleStrike: true,
      effect: "shimmer" as const,
      emboss: true,
      emphasisMark: "dot" as const,
      highlight: "yellow" as const,
      imprint: true,
      italic: true,
      italicCs: true,
      kerning: 8,
      outline: true,
      position: 4,
      scale: 120,
      shading: { pattern: "clear" as const, fill: { rgb: "00AA00" } },
      shadow: true,
      smallCaps: true,
      spacing: 20,
      strike: true,
      underline: { style: "single" as const },
      vertAlign: "superscript" as const,
    };
    const contextualStyles: StyleDefinitions = {
      docDefaults: { rPr: plainFormatting },
      styles: [
        { styleId: "PlainParagraph", type: "paragraph", name: "Plain Paragraph" },
        { styleId: "PlainCharacter", type: "character", name: "Plain Character" },
        {
          styleId: "ContextParagraph",
          type: "paragraph",
          name: "Context Paragraph",
          rPr: contextualFormatting,
        },
        {
          styleId: "ContextCharacter",
          type: "character",
          name: "Context Character",
          rPr: { color: { rgb: "AA0000" }, fontSize: 30, underline: { style: "double" } },
        },
      ],
    };
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "PlainParagraph" },
              content: [runText("From", { styleId: "PlainCharacter" })],
            },
            {
              type: "paragraph",
              formatting: { styleId: "ContextParagraph" },
              content: [runText("To", { styleId: "ContextCharacter" })],
            },
          ],
        },
        styles: contextualStyles,
      },
    };
    let state = createStyledEditorState(input, contextualStyles);
    expect(state.doc.child(1).attrs["defaultTextFormatting"]?.shading).toEqual({
      fill: { rgb: "00AA00" },
      pattern: "clear",
    });
    state = selectText(state, "From");
    const source = captureFormatMarks(state);
    state = selectText(state, "To");
    expect(
      applyFormatMarks(source)(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    let targetMarks;
    state.doc.descendants((node) => {
      if (node.isText && node.text === "To") {
        targetMarks = node.marks;
      }
    });
    expect(targetMarks).toBeDefined();
    expect(
      targetMarks?.find((mark) => mark.type.name === "runFormattingOverride")?.attrs,
    ).toMatchObject({ shading: { pattern: "nil" } });
    expect(
      marksToTextFormatting(targetMarks ?? [], { inheritedFormatting: contextualFormatting })
        .shading,
    ).toEqual({ pattern: "nil" });

    const saved = fromProseDoc(state.doc, input);
    expect(findRun(paragraphAt(saved, 1), "To").formatting?.shading).toEqual({ pattern: "nil" });
    const xml = serializeDocument(saved);
    expect(xml).toContain('<w:color w:val="auto"/>');
    expect(xml).toContain('<w:highlight w:val="none"/>');
    expect(xml).toContain('<w:shd w:val="nil"');
    expect(xml).toContain('<w:vertAlign w:val="baseline"/>');
    expect(xml).toContain('<w:effect w:val="none"/>');
    const reopened = reopenSerializedBody(saved);
    let reopenedState = createStyledEditorState(reopened, contextualStyles);
    reopenedState = selectText(reopenedState, "To");
    const target = captureFormatMarks(reopenedState);
    const targetFlowRun = toFlowBlocks(reopenedState.doc, { styles: contextualStyles })
      .flatMap((block) => (block.kind === "paragraph" ? block.runs : []))
      .find((run) => run.kind === "text" && run.text === "To");

    expect(target.effectiveFormatting).toEqual(source.effectiveFormatting);
    expect(targetFlowRun).toMatchObject({
      allCaps: false,
      bold: false,
      complexScriptBold: false,
      complexScriptItalic: false,
      emboss: false,
      forceComplexScript: false,
      horizontalScale: 100,
      imprint: false,
      italic: false,
      kerningMinPt: 0,
      positionPx: 0,
      smallCaps: false,
      strike: false,
      subscript: false,
      superscript: false,
      textOutline: false,
      textShadow: false,
      underline: false,
    });
    expect(targetFlowRun?.color).toBeUndefined();
    expect(targetFlowRun?.emphasisMark).toBeUndefined();
    expect(targetFlowRun?.highlight).toBeUndefined();
    expect(targetFlowRun?.letterSpacing).toBeUndefined();
    expect(targetFlowRun?.shading).toBeUndefined();
    expect(targetFlowRun?.textEffect).toBeUndefined();
    expect(findRun(paragraphAt(reopened, 1), "To").formatting?.styleId).toBe("PlainCharacter");
  });

  test("preserves a target revision through paint and save before accept or reject", () => {
    const revisionStyles: StyleDefinitions = {
      docDefaults: {
        rPr: {
          fontFamily: { asciiTheme: "minorHAnsi", hAnsiTheme: "minorHAnsi" },
          fontSize: 22,
        },
      },
      styles: [
        {
          styleId: "SourceCharacter",
          type: "character",
          name: "Source Character",
          rPr: { bold: true, color: { themeColor: "accent2", themeTint: "66" } },
        },
      ],
    };
    const input: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [runText("From", { styleId: "SourceCharacter" })],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "insertion",
                  info: { id: 7, author: "Reviewer", date: "2026-09-09T08:00:00Z" },
                  content: [runText("To")],
                },
              ],
            },
          ],
        },
        styles: revisionStyles,
      },
    };
    let state = createStyledEditorState(input, revisionStyles);
    state = selectText(state, "From");
    const source = captureFormatMarks(state);
    state = selectText(state, "To");
    expect(
      applyFormatMarks(source)(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const saved = fromProseDoc(state.doc, input);
    const reopened = reopenSerializedBody(saved);
    const pendingState = createStyledEditorState(reopened, revisionStyles);
    let pendingInsertionCount = 0;
    pendingState.doc.descendants((node) => {
      pendingInsertionCount += node.marks.filter((mark) => mark.type.name === "insertion").length;
    });
    expect(pendingInsertionCount).toBe(1);
    const savedTarget = paragraphAt(saved, 1).content.at(0);
    if (savedTarget?.type !== "insertion") {
      throw new Error("Expected pending insertion after paint and save");
    }
    const savedTargetRun = savedTarget.content.at(0);
    if (savedTargetRun?.type !== "run") {
      throw new Error("Expected formatted run inside pending insertion");
    }
    expect(savedTargetRun.formatting?.styleId).toBe("SourceCharacter");
    expect(savedTargetRun.formatting?.color).toBeUndefined();

    let accepted = pendingState;
    expect(
      acceptAllChanges()(accepted, (transaction) => {
        accepted = accepted.apply(transaction);
      }),
    ).toBe(true);
    expect(accepted.doc.textContent).toBe("FromTo");
    expect(JSON.stringify(accepted.doc.toJSON())).not.toContain("insertion");
    const acceptedTarget = captureFormatMarks(selectText(accepted, "To"));
    expect(acceptedTarget.effectiveFormatting.bold).toBe(source.effectiveFormatting.bold);
    expect(acceptedTarget.effectiveFormatting.color).toEqual(source.effectiveFormatting.color);

    let rejected = pendingState;
    expect(
      rejectAllChanges()(rejected, (transaction) => {
        rejected = rejected.apply(transaction);
      }),
    ).toBe(true);
    expect(rejected.doc.textContent).toBe("From");
    expect(JSON.stringify(rejected.doc.toJSON())).not.toContain("insertion");
  });

  test("hyperlink child runs keep their character style", () => {
    const input: Document = wrapParagraph({
      type: "paragraph",
      content: [
        {
          type: "hyperlink",
          href: "https://example.com",
          children: [runText("link", { styleId: "DefinedTerm" })],
        },
      ],
    });
    const pmDoc = toProseDoc(input, { styles });
    const out = fromProseDoc(pmDoc, input);
    const hyperlink = firstParagraph(out).content.find((content) => content.type === "hyperlink");
    if (hyperlink?.type !== "hyperlink") {
      throw new Error("Expected hyperlink");
    }
    const child = hyperlink.children.at(0);
    expect(child?.formatting?.styleId).toBe("DefinedTerm");
  });
});

describe("unknown and malformed style references", () => {
  test("unknown styleId round-trips verbatim with direct formatting intact", () => {
    const input = wrap(runText("Term", { styleId: "NoSuchStyle", bold: true }));
    const pmDoc = toProseDoc(input, { styles });
    const out = fromProseDoc(pmDoc, input);
    const run = findRun(firstParagraph(out), "Term");
    expect(run.formatting?.styleId).toBe("NoSuchStyle");
    expect(run.formatting?.bold).toBe(true);
  });

  test("unknown styleId resolves no formatting and does not crash", () => {
    const input = wrap(runText("Term", { styleId: "NoSuchStyle" }));
    const names = markNames(input, "Term");
    expect(names).toEqual(["characterStyle"]);
  });

  test("styleId round-trips without any style definitions at all", () => {
    const input = wrap(runText("Term", { styleId: "DefinedTerm" }));
    const pmDoc = toProseDoc(input);
    const out = fromProseDoc(pmDoc, input);
    expect(findRun(firstParagraph(out), "Term").formatting?.styleId).toBe("DefinedTerm");
  });

  test("basedOn cycle in style definitions terminates and round-trips", () => {
    const cyclicStyles: StyleDefinitions = {
      styles: [
        {
          styleId: "CycleA",
          type: "character",
          name: "Cycle A",
          basedOn: "CycleB",
          rPr: { bold: true },
        },
        {
          styleId: "CycleB",
          type: "character",
          name: "Cycle B",
          basedOn: "CycleA",
          rPr: { italic: true },
        },
      ],
    };
    const input = wrap(runText("Term", { styleId: "CycleA" }));
    const pmDoc = toProseDoc(input, { styles: cyclicStyles });
    const out = fromProseDoc(pmDoc, input);
    expect(findRun(firstParagraph(out), "Term").formatting?.styleId).toBe("CycleA");
  });
});

describe("character style under editing", () => {
  const styledState = (): EditorState => {
    const input = wrap(runText("Term", { styleId: "DefinedTerm" }));
    const pmDoc = toProseDoc(input, { styles });
    return EditorState.create({ doc: pmDoc });
  };

  test("typing in the middle of a styled run keeps the style", () => {
    const state = styledState();
    // Position 3 is between "Te" and "rm" (paragraph opens at 0, text at 1).
    const tr = state.tr.setSelection(TextSelection.create(state.doc, 3));
    tr.insertText("XY");
    const edited = state.apply(tr);
    const out = fromProseDoc(edited.doc, wrap());
    const run = findRun(firstParagraph(out), "TeXYrm");
    expect(run.formatting?.styleId).toBe("DefinedTerm");
  });

  test("splitting a styled run keeps the style on both halves", () => {
    const state = styledState();
    // Insert unmarked text in the middle: the styled run splits in two.
    const tr = state.tr.replaceWith(3, 3, schema.text("PLAIN"));
    const edited = state.apply(tr);
    const out = fromProseDoc(edited.doc, wrap());
    const runs = paragraphRuns(firstParagraph(out));
    expect(runs).toHaveLength(3);
    expect(findRun(firstParagraph(out), "Te").formatting?.styleId).toBe("DefinedTerm");
    expect(findRun(firstParagraph(out), "PLAIN").formatting).toBeUndefined();
    expect(findRun(firstParagraph(out), "rm").formatting?.styleId).toBe("DefinedTerm");
  });

  test("removing the characterStyle mark strips the reference but keeps visuals", () => {
    const state = styledState();
    const characterStyle = schema.marks["characterStyle"];
    if (!characterStyle) {
      throw new Error("Expected characterStyle mark type");
    }
    const tr = state.tr.removeMark(1, 5, characterStyle);
    const edited = state.apply(tr);
    const out = fromProseDoc(edited.doc, wrap());
    const run = findRun(firstParagraph(out), "Term");
    expect(run.formatting?.styleId).toBeUndefined();
    // The flattened rendering formatting is now genuinely direct.
    expect(run.formatting?.italic).toBe(true);
    expect(run.formatting?.color).toEqual({ rgb: "336699" });
  });

  test("toggling off a style-provided italic emits an explicit negative override", () => {
    const input = wrap(runText("Term", { styleId: "DefinedTerm" }));
    input.package.styles = styles;
    const state = EditorState.create({ doc: toProseDoc(input, { styles }) });
    const italic = schema.marks["italic"];
    if (!italic) {
      throw new Error("Expected italic mark type");
    }
    // Keep the character style reference but remove the italic the style
    // supplied, as the toolbar's italic toggle would.
    const tr = state.tr.removeMark(1, 5, italic);
    const edited = state.apply(tr);
    const out = fromProseDoc(edited.doc, input);
    const run = findRun(firstParagraph(out), "Term");
    // The style reference survives, but the removed italic must serialize as
    // an explicit negative so a consumer does not re-impose it from the style.
    expect(run.formatting?.styleId).toBe("DefinedTerm");
    expect(run.formatting?.italic).toBe(false);
    expect(run.formatting?.italicCs).toBeUndefined();
  });
});
