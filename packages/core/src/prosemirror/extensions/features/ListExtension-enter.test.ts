import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { computeListRendering, parseNumbering } from "../../../docx/numberingParser";
import { applyFolioAIEditOperations } from "../../../ai-edits/apply";
import { createFolioAIEditSnapshot } from "../../../ai-edits/snapshot";
import { toFlowBlocks } from "../../../layout-bridge/convert/toFlowBlocks";
import type { Document } from "../../../types/document";
import { fromProseDoc } from "../../conversion/fromProseDoc";
import { toProseDoc } from "../../conversion/toProseDoc";
import { LIST_RENDERING_ATTR_KEYS } from "../../listMarker";
import { createDocumentNumberingPlugin } from "../../plugins/documentNumbering";
import { schema } from "../../schema";
import { ListExtension, toggleNumberedList } from "./ListExtension";

const syntheticNumberedDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { numPr: { numId: 23, ilvl: 3 } },
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Synthetic numbered clause" }],
            },
          ],
          listRendering: {
            marker: "(a)",
            markerTemplate: "(%4)",
            level: 3,
            numId: 23,
            isBullet: false,
            numFmt: "lowerLetter",
            levelNumFmts: ["decimal", "decimal", "decimal", "lowerLetter"],
            abstractNumId: 15,
            startOverride: 1,
          },
        },
      ],
    },
  },
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MULTILEVEL_NUMBERING = parseNumbering(`
  <w:numbering ${W}>
    <w:abstractNum w:abstractNumId="41">
      <w:lvl w:ilvl="0">
        <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      </w:lvl>
      <w:lvl w:ilvl="1">
        <w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/>
        <w:lvlJc w:val="right"/><w:suff w:val="space"/>
        <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:i/></w:rPr>
      </w:lvl>
    </w:abstractNum>
    <w:num w:numId="23">
      <w:abstractNumId w:val="41"/>
      <w:lvlOverride w:ilvl="1"><w:startOverride w:val="3"/></w:lvlOverride>
    </w:num>
  </w:numbering>
`);

const multilevelDocument = (level: number): Document => {
  const listRendering = computeListRendering({ numId: 23, ilvl: level }, MULTILEVEL_NUMBERING);
  if (!listRendering) {
    return panic("Synthetic numbering did not contain its target level");
  }
  return {
    package: {
      numbering: MULTILEVEL_NUMBERING.definitions,
      document: {
        content: [
          {
            type: "paragraph",
            formatting: { numPr: { numId: 23, ilvl: level } },
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "Synthetic level transition" }],
              },
            ],
            listRendering,
          },
        ],
      },
    },
  };
};

const multilevelState = (level: number): EditorState =>
  EditorState.create({
    doc: toProseDoc(multilevelDocument(level)),
    plugins: [createDocumentNumberingPlugin(MULTILEVEL_NUMBERING.definitions)],
  });

const styleNumberedDocument = (): Document => ({
  package: {
    numbering: MULTILEVEL_NUMBERING.definitions,
    styles: {
      styles: [
        {
          styleId: "SyntheticClause",
          type: "paragraph",
          pPr: { numPr: { numId: 23, ilvl: 1 } },
        },
      ],
    },
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { styleId: "SyntheticClause" },
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Synthetic styled clause" }],
            },
          ],
        },
      ],
    },
  },
});

const listMarkers = (state: EditorState): string[] =>
  toFlowBlocks(state.doc).flatMap((block) =>
    block.kind === "paragraph" && block.attrs?.listMarker ? [block.attrs.listMarker] : [],
  );

describe("ListExtension Enter numbering", () => {
  test("advances an imported marker from its source template", () => {
    let state = EditorState.create({ doc: toProseDoc(syntheticNumberedDocument()) });
    const paragraph = state.doc.firstChild;
    if (!paragraph) {
      panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraph.nodeSize - 1)),
    );

    const enter = ListExtension().onSchemaReady({ schema }).keyboardShortcuts?.["Enter"];
    expect(enter).toBeDefined();
    if (!enter) {
      panic("List extension did not register Enter");
    }
    expect(
      enter(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(listMarkers(state)).toEqual(["(a)", "(b)"]);
    expect(state.selection.$from.parent).toBe(state.doc.lastChild);
    expect(state.selection.$from.parentOffset).toBe(0);
  });

  test("Backspace keeps a style-numbered list exit after the document model refreshes", () => {
    const document = styleNumberedDocument();
    let state = EditorState.create({
      doc: toProseDoc(document, { styles: document.package.styles }),
      plugins: [createDocumentNumberingPlugin(MULTILEVEL_NUMBERING.definitions)],
    });
    const runtime = ListExtension().onSchemaReady({ schema });
    const enter = runtime.keyboardShortcuts?.["Enter"];
    const backspace = runtime.keyboardShortcuts?.["Backspace"];
    if (!enter || !backspace) {
      return panic("List extension did not register Enter and Backspace");
    }
    const first = state.doc.firstChild;
    if (!first) {
      return panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, first.nodeSize - 1)));
    enter(state, (transaction) => {
      state = state.apply(transaction);
    });
    backspace(state, (transaction) => {
      state = state.apply(transaction);
    });

    const refreshedDocument = fromProseDoc(state.doc, document);
    const refreshedState = EditorState.create({
      doc: toProseDoc(refreshedDocument, { styles: refreshedDocument.package.styles }),
      plugins: [createDocumentNumberingPlugin(MULTILEVEL_NUMBERING.definitions)],
    });

    expect(listMarkers(refreshedState)).toHaveLength(1);
    expect(refreshedState.doc.lastChild?.attrs["numPr"]).toEqual({ numId: 0, ilvl: 1 });
  });

  test("does not retain an imported template after replacing the list", () => {
    let state = EditorState.create({ doc: toProseDoc(syntheticNumberedDocument()) });
    const runtime = ListExtension().onSchemaReady({ schema });
    const removeList = runtime.commands?.removeList?.();
    if (!removeList) {
      panic("List extension did not register removeList");
    }

    expect(
      removeList(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const attrs = state.doc.firstChild?.attrs;
    if (!attrs) {
      panic("Synthetic document did not contain its paragraph");
    }
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(attrs[key]).toBeNull();
    }

    expect(
      toggleNumberedList(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const paragraph = state.doc.firstChild;
    if (!paragraph) {
      panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraph.nodeSize - 1)),
    );

    const enter = runtime.keyboardShortcuts?.["Enter"];
    if (!enter) {
      panic("List extension did not register Enter");
    }
    expect(
      enter(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(listMarkers(state)).toEqual(["1.", "2."]);
  });

  test("clears every rendering attr when decreasing the first level exits the list", () => {
    let state = EditorState.create({
      doc: toProseDoc({
        ...syntheticNumberedDocument(),
        package: {
          document: {
            content: [
              {
                ...syntheticNumberedDocument().package.document.content[0],
                formatting: { numPr: { numId: 23, ilvl: 0 } },
              },
            ],
          },
        },
      }),
    });
    const decreaseListLevel = ListExtension()
      .onSchemaReady({ schema })
      .commands?.decreaseListLevel?.();
    if (!decreaseListLevel) {
      panic("List extension did not register decreaseListLevel");
    }

    expect(
      decreaseListLevel(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const attrs = state.doc.firstChild?.attrs;
    if (!attrs) {
      panic("Synthetic document did not contain its paragraph");
    }
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(attrs[key]).toBeNull();
    }
  });

  test.each([
    ["increaseListLevel command", 0, 1, "increaseListLevel"],
    ["Tab shortcut", 0, 1, "Tab"],
    ["decreaseListLevel command", 1, 0, "decreaseListLevel"],
    ["Shift-Tab shortcut", 1, 0, "Shift-Tab"],
  ] as const)("recomputes target-level attrs through the %s", (_name, source, target, action) => {
    let state = multilevelState(source);
    const runtime = ListExtension().onSchemaReady({ schema });
    const command =
      action === "Tab" || action === "Shift-Tab"
        ? runtime.keyboardShortcuts?.[action]
        : runtime.commands?.[action]?.();
    if (!command) {
      return panic("List extension did not register its level transition");
    }

    expect(
      command(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const attrs = state.doc.firstChild?.attrs;
    if (!attrs) {
      return panic("Synthetic document did not contain its paragraph");
    }
    expect(attrs["numPr"]).toEqual({ numId: 23, ilvl: target });
    if (target === 1) {
      expect(attrs).toMatchObject({
        listMarkerTemplate: "(%2)",
        listNumFmt: "lowerLetter",
        listMarkerAlignment: "right",
        listMarkerSuffix: "space",
        listMarkerFormatting: { italic: true },
        listStartOverride: 3,
        indentLeft: 1440,
        indentFirstLine: -360,
        hangingIndent: true,
      });
    } else {
      expect(attrs).toMatchObject({
        listMarkerTemplate: "%1.",
        listNumFmt: "decimal",
        listMarkerAlignment: "left",
        listMarkerSuffix: null,
        listStartOverride: null,
        indentLeft: 720,
        indentFirstLine: -360,
        hangingIndent: true,
      });
    }
  });

  test("Enter continues the recomputed target-level pattern after Tab", () => {
    let state = multilevelState(0);
    const runtime = ListExtension().onSchemaReady({ schema });
    const tab = runtime.keyboardShortcuts?.["Tab"];
    const enter = runtime.keyboardShortcuts?.["Enter"];
    if (!tab || !enter) {
      return panic("List extension did not register Tab and Enter");
    }
    expect(
      tab(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const paragraph = state.doc.firstChild;
    if (!paragraph) {
      return panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraph.nodeSize - 1)),
    );
    expect(
      enter(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(listMarkers(state)).toEqual(["(c)", "(d)"]);
  });

  test("paragraph operations use the same target-level projection", () => {
    const view = {
      state: multilevelState(0),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "synthetic-level-change",
          type: "setBlockParagraphProperties",
          blockId: "seq-0001",
          properties: { listLevel: 1 },
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.firstChild?.attrs).toMatchObject({
      numPr: { numId: 23, ilvl: 1 },
      listMarkerTemplate: "(%2)",
      listNumFmt: "lowerLetter",
      listStartOverride: 3,
      indentLeft: 1440,
    });
    expect(listMarkers(view.state)).toEqual(["(c)"]);
  });

  test("paragraph operations clear stale rendering even when numbering is already absent", () => {
    const initial = multilevelState(0);
    const paragraph = initial.doc.firstChild;
    if (!paragraph) {
      return panic("Synthetic document did not contain its paragraph");
    }
    const view = {
      state: initial.apply(
        initial.tr.setNodeMarkup(0, undefined, {
          ...paragraph.attrs,
          numPr: null,
        }),
      ),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "synthetic-list-removal",
          type: "setBlockParagraphProperties",
          blockId: "seq-0001",
          properties: { listLevel: null },
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.firstChild?.attrs["numPr"]).toBeNull();
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(view.state.doc.firstChild?.attrs[key]).toBeNull();
    }
  });

  test("paragraph operations canonicalize rendering when the requested level already matches", () => {
    const initial = multilevelState(0);
    const paragraph = initial.doc.firstChild;
    if (!paragraph) {
      return panic("Synthetic document did not contain its paragraph");
    }
    const view = {
      state: initial.apply(
        initial.tr.setNodeMarkup(0, undefined, {
          ...paragraph.attrs,
          listMarker: "(c)",
          listMarkerTemplate: "(%2)",
          listNumFmt: "lowerLetter",
        }),
      ),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "synthetic-list-canonicalization",
          type: "setBlockParagraphProperties",
          blockId: "seq-0001",
          properties: { listLevel: 0 },
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.firstChild?.attrs).toMatchObject({
      numPr: { numId: 23, ilvl: 0 },
      listMarker: "%1.",
      listMarkerTemplate: "%1.",
      listNumFmt: "decimal",
    });
    expect(listMarkers(view.state)).toEqual(["1."]);
  });

  test("paragraph insertion recomputes an explicitly requested list level", () => {
    const view = {
      state: multilevelState(0),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "synthetic-level-insert",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "Synthetic nested item",
          inheritFormatting: true,
          listLevel: 1,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(1).attrs).toMatchObject({
      numPr: { numId: 23, ilvl: 1 },
      listMarkerTemplate: "(%2)",
      listNumFmt: "lowerLetter",
      listStartOverride: 3,
      indentLeft: 1440,
    });
    expect(listMarkers(view.state)).toEqual(["1.", "(c)"]);
  });
});
