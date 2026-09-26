/**
 * Every list gesture × every package shape × both editing modes, through a
 * save and a reparse.
 *
 * A list command's result is only real once it is saved: the paragraph can
 * paint as a list item while the model references a numbering instance the
 * package does not define, and the first sign is a save that throws (#1091).
 * It can also join a list it has no business joining, which the save accepts
 * and a reader only notices as renumbered clauses (#1092). So each cell here
 * applies one gesture the way the editor applies it (the command, or the typed
 * marker fed through the text-input funnel with the editor's plugin order),
 * saves with `repackDocx`, parses the bytes back, and checks the invariants
 * that hold for every list in every package:
 *
 * - the save succeeds and every `w:numPr` names an instance the package defines;
 * - the paragraph is a list item of the requested kind, showing the requested
 *   first marker (a new list; the target never touches another list);
 * - every other paragraph keeps its label (no list, heading numbering included,
 *   is renumbered by a gesture elsewhere);
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
import { FolioDocxReviewer } from "../../../ai-edits/headless";
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
import { continueNumbering, restartNumbering, setNumberingValue } from "../../listNumbering";
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

type Label = { text: string; label: string };

const labelsOf = async (bytes: ArrayBuffer): Promise<Label[]> =>
  (await FolioDocxReviewer.fromBuffer(bytes))
    .getContent()
    .map(({ text, displayLabel }) => ({ text, label: displayLabel ?? "" }));

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
            const before = await labelsOf(await repackDocx(document));
            const editor = editorFor(document, mode);

            applyGesture(editor, gesture);

            for (const saved of await savesOf(editor, document)) {
              await expectListInvariants({ saved, before, gesture, mode });
            }
          });
        }
      });
    }
  }
});

type ListInvariantsOptions = {
  saved: ArrayBuffer;
  before: readonly Label[];
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
  before,
  gesture,
  mode,
}: ListInvariantsOptions): Promise<void> => {
  const reparsed = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
  expectEveryReferenceDefined(reparsed);

  const target = targetParagraph(reparsed);
  expect(target.listRendering?.isBullet ?? false).toBe(gesture.expected.kind === "bullet");

  const after = await labelsOf(saved);
  expect(after.map(({ text }) => text)).toEqual(before.map(({ text }) => text));
  expect(after.find(({ text }) => text === TARGET)?.label).toBe(gesture.expected.label);
  expect(after.filter(({ text }) => text !== TARGET)).toEqual(
    before.filter(({ text }) => text !== TARGET),
  );

  // A new list: no other paragraph shares the target's instance.
  const targetNumId = paragraphNumberingReferenceId(target.formatting?.numPr);
  const sharing = paragraphsOf(reparsed.package.document.content).filter(
    (paragraph) =>
      paragraph !== target &&
      paragraphNumberingReferenceId(paragraph.formatting?.numPr) === targetNumId,
  );
  expect(sharing).toHaveLength(0);

  // Suggesting records exactly one change whose previous state is
  // the plain paragraph, so a reject restores it.
  const changes = target.propertyChanges ?? [];
  expect(changes).toHaveLength(mode === "suggesting" ? 1 : 0);
  expect(
    paragraphNumberingReferenceId(changes.at(0)?.previousFormatting?.numPr ?? undefined),
  ).toBeUndefined();
};

// ============================================================================
// WHICH LIST A GESTURE JOINS
// ============================================================================

type Session = {
  document: Document;
  editor: Editor;
};

const sessionOf = async (model: Document, mode: EditingMode = "editing"): Promise<Session> => {
  const document = await packageOf(model);
  return { document, editor: editorFor(document, mode) };
};

const caretAt = (editor: Editor, text: string): void => {
  const start = paragraphStart(editor.state, text);
  editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, start)));
};

const run = (editor: Editor, text: string, command: Command): boolean => {
  caretAt(editor, text);
  return command(editor.state, editor.dispatch);
};

const typeAt = (editor: Editor, text: string, marker: string): void => {
  caretAt(editor, text);
  typeText(editor, marker);
};

/** Save, check every reference is defined, and read the labels back. */
const savedLabels = async ({ document, editor }: Session): Promise<string[]> => {
  const saved = await repackDocx(fromProseDoc(editor.state.doc, document));
  expectEveryReferenceDefined(
    await parseDocx(saved, { preloadFonts: false, detectVariables: false }),
  );
  return (await labelsOf(saved)).map(({ text, label }) => `${label} ${text}`.trim());
};

type Gesturer = (editor: Editor, text: string) => void;

const clickNumbered: Gesturer = (editor, text) => {
  run(editor, text, toggleNumberedList);
};

const typeOne: Gesturer = (editor, text) => {
  typeAt(editor, text, "1. ");
};

describe("which list a gesture joins", () => {
  const secondList =
    "1. Alpha\n2. Beta\n\nUnrelated prose between the lists.\n\nNew list one\n\nNew list two";

  test.each([
    ["Numbered List", clickNumbered],
    ["typed 1.", typeOne],
  ] as const)(
    "%s after prose starts a new list, and the next paragraph continues it",
    async (_name, gesture) => {
      const session = await sessionOf(fromMarkdown(secondList));
      gesture(session.editor, "New list one");
      gesture(session.editor, "New list two");

      expect(await savedLabels(session)).toEqual([
        "1. Alpha",
        "2. Beta",
        "Unrelated prose between the lists.",
        "1. New list one",
        "2. New list two",
      ]);
    },
  );

  test("a paragraph directly under a list of the requested kind continues it", async () => {
    const session = await sessionOf(fromMarkdown("1. Alpha\n2. Beta\n\nGamma"));
    run(session.editor, "Gamma", toggleNumberedList);

    expect(await savedLabels(session)).toEqual(["1. Alpha", "2. Beta", "3. Gamma"]);
  });

  test("a paragraph directly under a list of the other kind starts its own", async () => {
    const session = await sessionOf(fromMarkdown("1. Alpha\n2. Beta\n\nGamma"));
    run(session.editor, "Gamma", toggleBulletList);

    expect(await savedLabels(session)).toEqual(["1. Alpha", "2. Beta", "• Gamma"]);
  });

  test("a body paragraph never joins heading numbering its style supplies", async () => {
    const session = await sessionOf(styleNumberedHeadings());
    run(session.editor, TARGET, toggleNumberedList);
    typeAt(session.editor, "Payment is due in ten days.", "1. ");

    const labels = await savedLabels(session);
    expect(labels).toContain("1. Scope");
    expect(labels).toContain("2. Payment");
    expect(labels).toContain(`1. ${TARGET}`);
    expect(labels).toContain("1. Payment is due in ten days.");
  });
});

// ============================================================================
// RESTART / CONTINUE / SET VALUE
// ============================================================================

describe("restart, continue and set numbering value", () => {
  const threeItems = "1. Alpha\n2. Beta\n3. Gamma";

  for (const mode of EDITING_MODES) {
    test(`restart at 1, ${mode}`, async () => {
      const session = await sessionOf(fromMarkdown(threeItems), mode);
      expect(run(session.editor, "Beta", restartNumbering)).toBe(true);

      expect(await savedLabels(session)).toEqual(["1. Alpha", "1. Beta", "2. Gamma"]);
    });

    test(`set numbering value, ${mode}`, async () => {
      const session = await sessionOf(fromMarkdown(threeItems), mode);
      expect(run(session.editor, "Beta", setNumberingValue(5))).toBe(true);

      expect(await savedLabels(session)).toEqual(["1. Alpha", "5. Beta", "6. Gamma"]);
    });
  }

  for (const mode of ["editing"] as const) {
    test(`continue undoes a restart, ${mode}`, async () => {
      const session = await sessionOf(fromMarkdown(threeItems), mode);
      run(session.editor, "Beta", restartNumbering);
      expect(run(session.editor, "Beta", continueNumbering)).toBe(true);

      expect(await savedLabels(session)).toEqual(["1. Alpha", "2. Beta", "3. Gamma"]);
    });
  }

  test("continue carries a second list on from the first", async () => {
    const session = await sessionOf(fromMarkdown("1. Alpha\n2. Beta\n\nProse.\n\nGamma\n\nDelta"));
    run(session.editor, "Gamma", toggleNumberedList);
    run(session.editor, "Delta", toggleNumberedList);
    expect(run(session.editor, "Gamma", continueNumbering)).toBe(true);

    expect(await savedLabels(session)).toEqual([
      "1. Alpha",
      "2. Beta",
      "Prose.",
      "3. Gamma",
      "4. Delta",
    ]);
  });

  test("are not applicable outside a list, or with no earlier list to continue", async () => {
    const { editor } = await sessionOf(fromMarkdown("Plain.\n\n1. Alpha"));
    caretAt(editor, "Plain.");
    expect(restartNumbering(editor.state)).toBe(false);
    expect(setNumberingValue(3)(editor.state)).toBe(false);
    caretAt(editor, "Alpha");
    expect(continueNumbering(editor.state)).toBe(false);
    expect(restartNumbering(editor.state)).toBe(true);
  });
});
