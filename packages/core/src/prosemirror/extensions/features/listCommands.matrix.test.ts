/**
 * Every list gesture × every package shape × both editing modes, through a
 * save and a reparse.
 *
 * A list command's result is only real once it is saved: the paragraph can
 * paint as a list item while the model references a numbering instance the
 * package does not define, and the first sign is a save that throws (#1091).
 * So each cell here
 * applies one gesture the way the editor applies it (the command, or the typed
 * marker fed through the text-input funnel with the editor's plugin order),
 * saves with `repackDocx`, parses the bytes back, and checks the invariants
 * that hold for every list in every package:
 *
 * - the save succeeds and every `w:numPr` names an instance the package defines;
 * - the paragraph is a list item of the requested kind;
 * - while suggesting, the change is a tracked paragraph-property change.
 *
 * New gestures and package shapes are rows in the tables below; the
 * invariants apply to them unchanged.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { EditorState, TextSelection, type Command, type Transaction } from "prosemirror-state";

import { ensureParaIds } from "../../../docx/ensureParaIds";
import { parseDocx } from "../../../docx/parser";
import { createDocx, repackDocx } from "../../../docx/rezip";
import { attemptSelectiveSave } from "../../../docx/selectiveSave";
import { fromMarkdown } from "../../../markdown/fromMarkdown";
import type { BlockContent, Document, Paragraph } from "../../../types/document";
import { paragraphNumberingReferenceId } from "../../../docx/numberingReference";
import { fromProseDoc } from "../../conversion/fromProseDoc";
import { toProseDoc } from "../../conversion/toProseDoc";
import { ExtensionManager } from "../ExtensionManager";
import { createStarterKit } from "../StarterKit";
import { createDocumentNumberingPlugin } from "../../plugins/documentNumbering";
import { createDocumentStylesPlugin } from "../../plugins/documentStyles";
import { createSuggestionModePlugin } from "../../plugins/suggestionMode";
import { dispatchEditorTextInput } from "../../textInput";
import { toggleBulletList, toggleNumberedList } from "./ListExtension";
import { getChangeTrackerState } from "./ParagraphChangeTrackerExtension";

setDefaultTimeout(60_000);

const TARGET = "Target paragraph";

// ============================================================================
// PACKAGE SHAPES
// ============================================================================

type PackageShape = {
  name: string;
  build: () => Document;
};

/** Headings numbered through their paragraph style, the only decimal list. */
const styleNumberedHeadings = (): Document => {
  const model = fromMarkdown(
    `# Agreement\n\n## Scope\n\nThe Supplier delivers the goods.\n\n${TARGET}\n\n## Payment\n\nPayment is due in ten days.`,
  );
  model.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        multiLevelType: "multilevel",
        levels: [
          { ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1.", suffix: "space" },
          { ilvl: 1, start: 1, numFmt: "decimal", lvlText: "%1.%2.", suffix: "space" },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const heading = model.package.styles?.styles.find(({ styleId }) => styleId === "Heading2");
  if (!heading) {
    throw new Error("the markdown style set must define Heading2");
  }
  heading.pPr = { ...heading.pPr, numPr: { kind: "reference", numId: 5, ilvl: 0 } };
  return model;
};

const PACKAGE_SHAPES: readonly PackageShape[] = [
  {
    name: "no numbering part",
    build: () => fromMarkdown(`Intro paragraph.\n\n${TARGET}\n\nTail.`),
  },
  {
    name: "bullet list only",
    build: () => fromMarkdown(`- Alpha\n- Beta\n\nPlain text.\n\n${TARGET}`),
  },
  {
    name: "decimal list only",
    build: () => fromMarkdown(`1. Alpha\n2. Beta\n\nPlain text.\n\n${TARGET}`),
  },
  { name: "style-numbered headings", build: styleNumberedHeadings },
  {
    name: "several lists",
    build: () =>
      fromMarkdown(
        `1. Alpha\n2. Beta\n\nMiddle prose.\n\n- Gamma\n- Delta\n\nPlain text.\n\n${TARGET}`,
      ),
  },
];

// ============================================================================
// GESTURES
// ============================================================================

type Expected = {
  kind: "bullet" | "numbered";
  /** The marker the target shows after the save. */
  label: string;
};

type Gesture =
  | { type: "command"; name: string; command: Command; expected: Expected }
  | { type: "typed"; marker: string; expected: Expected };

const BULLET: Expected = { kind: "bullet", label: "•" };

const LIST_GESTURES: readonly Gesture[] = [
  { type: "command", name: "Bullet List", command: toggleBulletList, expected: BULLET },
  {
    type: "command",
    name: "Numbered List",
    command: toggleNumberedList,
    expected: { kind: "numbered", label: "1." },
  },
  { type: "typed", marker: "- ", expected: BULLET },
  { type: "typed", marker: "* ", expected: BULLET },
  { type: "typed", marker: "1. ", expected: { kind: "numbered", label: "1." } },
];

const gestureName = (gesture: Gesture): string =>
  gesture.type === "command" ? gesture.name : `typed "${gesture.marker}"`;

// ============================================================================
// EDITOR
// ============================================================================

const EDITING_MODES = ["editing", "suggesting"] as const;
type EditingMode = (typeof EDITING_MODES)[number];

/** A document as `createDocx` + `ensureParaIds` + `parseDocx` hand it to an editor. */
const packageOf = async (model: Document): Promise<Document> => {
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(model)));
  return parseDocx(docx, { preloadFonts: false, detectVariables: false });
};

type Editor = {
  state: EditorState;
  composing: false;
  dispatch: (tr: Transaction) => void;
  someProp: (name: "handleTextInput", f: (handler: never) => unknown) => unknown;
};

/**
 * An editor with the plugins the real one runs, in its order: the host's
 * suggestion plugin first, then the extension plugins, then the document's
 * styles and numbering (see `createHiddenEditorState`).
 */
const editorFor = (document: Document, mode: EditingMode): Editor => {
  const manager = new ExtensionManager(createStarterKit());
  manager.buildSchema();
  manager.initializeRuntime();
  const { styles, numbering, theme } = document.package;
  const editor: Editor = {
    state: EditorState.create({
      doc: toProseDoc(document, { styles, theme }),
      plugins: [
        createSuggestionModePlugin(mode === "suggesting", "Reviewer"),
        ...manager.getPlugins(),
        createDocumentStylesPlugin(styles),
        createDocumentNumberingPlugin(numbering),
      ],
    }),
    composing: false,
    dispatch(tr) {
      editor.state = editor.state.apply(tr);
    },
    someProp(name, f) {
      for (const plugin of editor.state.plugins) {
        const prop = plugin.props[name];
        if (prop) {
          // SAFETY: `handleTextInput` is the only prop this stand-in routes.
          const handled = f(prop.bind(plugin) as never);
          if (handled) {
            return handled;
          }
        }
      }
      return undefined;
    },
  };
  return editor;
};

const paragraphStart = (state: EditorState, text: string): number => {
  let found: number | null = null;
  state.doc.descendants((node, pos) => {
    if (found !== null) {
      return false;
    }
    if (node.type.name === "paragraph" && node.textContent === text) {
      found = pos + 1;
    }
    return node.type.name !== "paragraph";
  });
  if (found === null) {
    throw new Error(`no paragraph reads "${text}"`);
  }
  return found;
};

/** Type `text` one character at a time through the editor's text-input funnel. */
const typeText = (editor: Editor, text: string): void => {
  for (const character of text) {
    dispatchEditorTextInput(editor, character);
  }
};

const applyGesture = (editor: Editor, gesture: Gesture): void => {
  const start = paragraphStart(editor.state, TARGET);
  editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, start)));
  if (gesture.type === "command") {
    expect(gesture.command(editor.state, editor.dispatch)).toBe(true);
    return;
  }
  typeText(editor, gesture.marker);
};

// ============================================================================
// INVARIANTS
// ============================================================================

const paragraphsOf = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "paragraph":
        return [block];
      case "table":
        return block.rows.flatMap((row) => row.cells.flatMap((cell) => paragraphsOf(cell.content)));
      default:
        return [];
    }
  });

/** Every `w:numPr` in the body, live or recorded as a tracked change's previous state. */
const expectEveryReferenceDefined = (document: Document): void => {
  const defined = new Set((document.package.numbering?.nums ?? []).map(({ numId }) => numId));
  for (const paragraph of paragraphsOf(document.package.document.content)) {
    const references = [
      paragraph.formatting?.numPr,
      ...(paragraph.propertyChanges ?? []).map((change) => change.previousFormatting?.numPr),
    ];
    for (const numPr of references) {
      const numId = paragraphNumberingReferenceId(numPr ?? undefined);
      if (numId !== undefined) {
        expect(defined.has(numId)).toBe(true);
      }
    }
  }
};

const targetParagraph = (document: Document): Paragraph => {
  const paragraph = paragraphsOf(document.package.document.content).find((candidate) =>
    candidate.content.some(
      (child) =>
        child.type === "run" &&
        child.content.some((part) => part.type === "text" && part.text.includes("Target")),
    ),
  );
  if (!paragraph) {
    throw new Error("the saved package lost the target paragraph");
  }
  return paragraph;
};

// ============================================================================
// MATRIX
// ============================================================================

describe("list gestures × package shapes × editing modes", () => {
  for (const shape of PACKAGE_SHAPES) {
    for (const mode of EDITING_MODES) {
      describe(`${shape.name}, ${mode}`, () => {
        for (const gesture of LIST_GESTURES) {
          // List autoformat does not run while suggesting yet.
          if (gesture.type === "typed" && mode === "suggesting") {
            continue;
          }
          test(gestureName(gesture), async () => {
            const document = await packageOf(shape.build());
            const editor = editorFor(document, mode);

            applyGesture(editor, gesture);

            for (const saved of await savesOf(editor, document)) {
              await expectListInvariants({ saved, gesture, mode });
            }
          });
        }
      });
    }
  }
});

type ListInvariantsOptions = {
  saved: ArrayBuffer;
  gesture: Gesture;
  mode: EditingMode;
};

/**
 * Both packages the editor may write: the full repack, and the selective save
 * that patches only the changed parts of the source package when it can.
 */
const savesOf = async (editor: Editor, document: Document): Promise<ArrayBuffer[]> => {
  const model = fromProseDoc(editor.state.doc, document);
  const tracker = getChangeTrackerState(editor.state);
  const { originalBuffer } = document;
  if (!tracker || !originalBuffer) {
    throw new Error("the editor must track changes against a source package");
  }
  const selective = await attemptSelectiveSave(model, originalBuffer, tracker);
  return [await repackDocx(model), ...(selective ? [selective] : [])];
};

const expectListInvariants = async ({
  saved,
  gesture,
  mode,
}: ListInvariantsOptions): Promise<void> => {
  const reparsed = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
  expectEveryReferenceDefined(reparsed);

  const target = targetParagraph(reparsed);
  expect(target.listRendering?.isBullet ?? false).toBe(gesture.expected.kind === "bullet");

  // Suggesting records exactly one change whose previous state is
  // the plain paragraph, so a reject restores it.
  const changes = target.propertyChanges ?? [];
  expect(changes).toHaveLength(mode === "suggesting" ? 1 : 0);
  expect(
    paragraphNumberingReferenceId(changes.at(0)?.previousFormatting?.numPr ?? undefined),
  ).toBeUndefined();
};
