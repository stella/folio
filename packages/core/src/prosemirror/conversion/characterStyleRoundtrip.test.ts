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
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import type { Document, Paragraph, Run, StyleDefinitions } from "../../types/document";
import { schema } from "../schema";
import { applyFormatMarks, captureFormatMarks } from "../commands/formatPainter";
import { fromProseDoc } from "./fromProseDoc";
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
        bold: true,
        fontSizePt: 15,
        color: "#FF0000",
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
    let state = EditorState.create({ doc: toProseDoc(input, { styles }) });
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
    let state = EditorState.create({ doc: toProseDoc(input, { styles }) });
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

  test("painting independent complex-script style toggles saves direct offs", () => {
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
    let state = EditorState.create({ doc: toProseDoc(input, { styles }) });
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
