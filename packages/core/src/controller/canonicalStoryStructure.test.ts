import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { OpStory } from "@stll/docx-core/ops";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import type { CanonicalRefusalGround } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

type StoryCase = {
  name: string;
  shape: string;
  story: OpStory;
  /** The word the caret goes after, and the text it splits. */
  text: string;
};

const CASES: readonly StoryCase[] = [
  {
    name: "header",
    shape: "header-footer",
    story: { kind: "header", rId: "rIdHeader1" },
    text: "Header text",
  },
  {
    name: "footer",
    shape: "header-footer",
    story: { kind: "footer", rId: "rIdFooter1" },
    text: "Footer text",
  },
  {
    name: "footnote",
    shape: "notes",
    story: { kind: "footnote", id: 1 },
    text: "The footnote text.",
  },
];

const parse = async (shape: string) =>
  parseDocx(await shapeArrayBuffer(shape), {
    preloadFonts: false,
    detectVariables: false,
  });

/** Each paragraph's text in a story projection. */
const paragraphTexts = (doc: PMNode): string[] => {
  const texts: string[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    texts.push(node.textContent);
    return false;
  });
  return texts;
};

/** The document position just before `needle` in the first text holding it. */
const positionBefore = (doc: PMNode, needle: string): number => {
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    const index = node.isText ? (node.text ?? "").indexOf(needle) : -1;
    if (index >= 0) found = position + index;
    return found === undefined;
  });
  return found ?? panic(`The story holds no "${needle}".`);
};

const mount = async ({ shape, story }: StoryCase, suggesting = false) => {
  const releaseDom = acquireHarnessDom();
  const bodyHost = document.createElement("div");
  const storyHost = document.createElement("div");
  document.body.append(bodyHost, storyHost);
  const source: Document = await parse(shape);
  const refused: { message: string; ground: CanonicalRefusalGround }[] = [];
  const onSessionRefusal = (message: string, ground: CanonicalRefusalGround) =>
    refused.push({ message, ground });
  const deps = {
    getHost: () => bodyHost,
    getDocument: () => source,
    getStyles: () => source.package.styles ?? null,
    getEmbeddedFontFamilyMap: () => undefined,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getReadOnly: () => false,
    getDocumentIdentity: () => `story-structure-${shape}`,
    getDocumentContext: () => source,
    getEditingMode: () => (suggesting ? "suggesting" : "editing"),
    getSuggestionAuthor: () => "Story author",
    onTransaction: () => undefined,
    onCanonicalCompositionLayout: () => undefined,
    onSelectionChange: () => undefined,
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => undefined,
    onEditorViewReady: () => undefined,
    onEditorViewDestroy: () => undefined,
    onRemoteSelectionsChange: () => undefined,
    onSessionRefusal,
  } satisfies HiddenEditorManagerDeps;
  const manager = createHiddenEditorManager(deps);
  const storyDeps = {
    getHost: () => storyHost,
    getDocument: () => source,
    getCanonicalApi: () => manager.api,
    readCanonicalDocument: () => manager.readCanonicalDocument(),
    getStyles: () => source.package.styles ?? null,
    getTheme: () => null,
    onSessionRefusal,
  };
  const parts = createHeaderFooterEditorManager(storyDeps);
  const notes = createNoteEditorManager(storyDeps);
  manager.ensureView();
  parts.sync();
  const storyView = (): EditorView | null => {
    if (story.kind === "header" || story.kind === "footer")
      return parts.getView(story.rId);
    if (story.kind === "footnote" || story.kind === "endnote")
      return notes.activate({ kind: story.kind, noteId: story.id });
    return null;
  };
  const view = storyView() ?? panic("The story editor did not mount.");
  const projection = () =>
    manager.api.getCanonicalStoryProjection(story) ??
    panic("The story has no projection.");
  const select = (from: number, to = from) =>
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, from, to),
      ),
    );
  const dispose = () => {
    parts.destroy();
    notes.destroy();
    manager.destroyView();
    bodyHost.remove();
    storyHost.remove();
    releaseDom();
  };
  return { manager, view, projection, select, refused, dispose };
};

const pressKey = (view: EditorView, key: string, shiftKey = false) =>
  view.dom.dispatchEvent(
    new KeyboardEvent("keydown", {
      key,
      shiftKey,
      bubbles: true,
      cancelable: true,
    }),
  );
const pressEnter = (view: EditorView) => pressKey(view, "Enter");

describe("story editor Enter and cut", () => {
  for (const storyCase of CASES) {
    const { name, story, text } = storyCase;
    // The caret goes before the last word, with its space.
    const first = text.slice(0, text.lastIndexOf(" "));
    const rest = text.slice(text.lastIndexOf(" "));
    // A note's text follows its reference mark and a space.
    const endsWith = (suffix: string) => (texts: string[]) =>
      texts.some((paragraph) => paragraph.endsWith(suffix));

    test(`Enter splits the ${name} paragraph, undo restores it in one step, and the save keeps the split`, async () => {
      const editor = await mount(storyCase);
      try {
        const before = editor.manager.api.getCanonicalDocument();
        editor.select(positionBefore(editor.view.state.doc, rest));
        pressEnter(editor.view);
        expect(editor.refused).toEqual([]);
        const split = paragraphTexts(editor.projection());
        expect(endsWith(first)(split)).toBe(true);
        expect(split).toContain(rest);
        expect(endsWith(text)(split)).toBe(false);
        expect(editor.view.state.doc.eq(editor.projection())).toBe(true);
        const snapshot =
          editor.manager.api.captureCanonicalSave() ??
          panic("The session has no save snapshot.");
        const saved = await serializeCanonicalSave({
          snapshot,
          options: { mode: "full" },
        });
        const reopened = createCanonicalSession(
          await parseDocx(saved.buffer, {
            preloadFonts: false,
            detectVariables: false,
          }),
        ).unwrap();
        expect(
          paragraphTexts(reopened.projectStory(story).unwrap().doc),
        ).toEqual(paragraphTexts(editor.projection()));
        expect(
          editor.manager.api.applyCanonicalStoryHistory({
            view: editor.view,
            story,
            direction: "undo",
          }),
        ).toBe(true);
        expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
        expect(endsWith(text)(paragraphTexts(editor.projection()))).toBe(true);
        expect(editor.manager.api.canUndo()).toBe(false);
      } finally {
        editor.dispose();
      }
    });

    test(`cut in the ${name} removes the selection and writes it to the clipboard, and undo restores it`, async () => {
      const editor = await mount(storyCase);
      try {
        const before = editor.manager.api.getCanonicalDocument();
        const from = positionBefore(editor.view.state.doc, rest);
        editor.select(from, from + rest.length);
        const clipboardData = new DataTransfer();
        const event = new ClipboardEvent("cut", {
          bubbles: true,
          cancelable: true,
          clipboardData,
        });
        editor.view.dom.dispatchEvent(event);
        expect(editor.refused).toEqual([]);
        expect(event.defaultPrevented).toBe(true);
        expect(clipboardData.getData("text/plain")).toBe(rest);
        const cut = paragraphTexts(editor.projection());
        expect(endsWith(first)(cut)).toBe(true);
        expect(endsWith(text)(cut)).toBe(false);
        expect(editor.manager.api.canUndo()).toBe(true);
        expect(
          editor.manager.api.applyCanonicalStoryHistory({
            view: editor.view,
            story,
            direction: "undo",
          }),
        ).toBe(true);
        expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
      } finally {
        editor.dispose();
      }
    });
  }

  test("Backspace at the start of a split header paragraph joins it back, and Shift+Enter breaks the line", async () => {
    const storyCase = CASES[0] ?? panic("The header case is declared.");
    const editor = await mount(storyCase);
    try {
      const at = positionBefore(editor.view.state.doc, " text");
      editor.select(at);
      pressEnter(editor.view);
      // The caret now starts the second paragraph.
      expect(editor.view.state.selection.$from.parentOffset).toBe(0);
      pressKey(editor.view, "Backspace");
      expect(editor.refused).toEqual([]);
      expect(paragraphTexts(editor.projection())).toEqual(["Header text"]);
      editor.select(positionBefore(editor.view.state.doc, " text"));
      pressKey(editor.view, "Enter", true);
      expect(editor.refused).toEqual([]);
      const breaks: string[] = [];
      editor.projection().descendants((node) => {
        if (node.type.name === "hardBreak") breaks.push(node.type.name);
      });
      expect(paragraphTexts(editor.projection())).toHaveLength(1);
      expect(breaks).toHaveLength(1);
    } finally {
      editor.dispose();
    }
  });

  test("Enter in a header while suggesting records the new paragraph mark", async () => {
    const storyCase = CASES[0] ?? panic("The header case is declared.");
    const editor = await mount(storyCase, true);
    try {
      editor.select(positionBefore(editor.view.state.doc, " text"));
      pressEnter(editor.view);
      expect(editor.refused).toEqual([]);
      const header = editor.manager.api
        .getCanonicalDocument()
        ?.package.headers?.get("rIdHeader1");
      const paragraphs = (header?.content ?? []).filter(
        (block) => block.type === "paragraph",
      );
      expect(paragraphs).toHaveLength(2);
      expect(paragraphs.at(0)?.pPrMark?.kind).toBe("ins");
    } finally {
      editor.dispose();
    }
  });

  test("a page break stays out of a footnote", async () => {
    const storyCase = CASES[2] ?? panic("The footnote case is declared.");
    const editor = await mount(storyCase);
    try {
      const before = editor.manager.api.getCanonicalDocument();
      editor.select(positionBefore(editor.view.state.doc, " text"));
      expect(
        editor.manager.api.applyCanonicalStoryStructure({
          view: editor.view,
          story: storyCase.story,
          type: "pageBreak",
        }),
      ).toBe(false);
      expect(editor.refused).toEqual([
        {
          message: "A page break goes in the main story.",
          ground: { condition: "invalidRequest" },
        },
      ]);
      expect(editor.manager.api.getCanonicalDocument()).toEqual(before);
    } finally {
      editor.dispose();
    }
  });
});
