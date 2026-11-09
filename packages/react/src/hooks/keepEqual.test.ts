import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { keepEqual, structurallyEqual } from "./keepEqual";

describe("keepEqual", () => {
  test("keeps the current value for an equal one, from a value or an updater", () => {
    const current = [{ x: 1 }];
    expect(keepEqual(current, [{ x: 1 }])).toBe(current);
    expect(keepEqual(current, () => [{ x: 1 }])).toBe(current);
    const next = [{ x: 2 }];
    expect(keepEqual(current, next)).toBe(next);
    expect(keepEqual(current, (previous) => [...previous, { x: 3 }])).toEqual([
      { x: 1 },
      { x: 3 },
    ]);
  });
});

describe("structurallyEqual", () => {
  test("compares data by content and other objects by identity", () => {
    expect(structurallyEqual([], [])).toBe(true);
    expect(
      structurallyEqual(
        [{ x: 1, y: 2, nested: [3] }],
        [{ x: 1, y: 2, nested: [3] }],
      ),
    ).toBe(true);
    expect(structurallyEqual({ a: 1 }, { a: 1, b: undefined })).toBe(false);
    expect(structurallyEqual([1, 2], [2, 1])).toBe(false);
    expect(structurallyEqual(new Map([["a", 1]]), new Map([["a", 1]]))).toBe(
      true,
    );
    expect(structurallyEqual(new Set([1]), new Set([2]))).toBe(false);
    expect(structurallyEqual(Number.NaN, Number.NaN)).toBe(true);
    expect(structurallyEqual(null, {})).toBe(false);
    class Box {
      constructor(readonly value: number) {}
    }
    expect(structurallyEqual(new Box(1), new Box(1))).toBe(false);
    const same = new Box(1);
    expect(structurallyEqual({ box: same }, { box: same })).toBe(true);
  });
});

/**
 * State the editor host writes while the user types must not commit when the
 * value is unchanged: every commit with the editor focused reads its whole
 * DOM. Object-valued state there is declared with `useEqualState`. A plain
 * `useState` with a setter is listed here with why a raw write is fine.
 */
const RAW_STATE: Record<string, Record<string, string>> = {
  "components/DocxEditor.tsx": {
    setOutlineDepthState: "a user choice",
    setState: "its updaters return the previous state when nothing changed",
    setTablePropsOpen: "boolean",
    setFootnotePropsOpen: "boolean",
    setShowInsertSymbol: "boolean",
    setNumberingValueDialogOpen: "boolean",
    setRulerVisible: "boolean",
    setShowOutline: "boolean",
    setPreviousShowOutlineProp: "boolean",
    setBodyViewEpoch: "number",
    setBodyEditorView: "set when the body editor view mounts",
    setEditorController: "set when an editor mounts",
    setEditorContentElement: "set when the content element mounts",
    setScrollContainer: "set when the scroll container mounts",
    setWidestLaidOutPageWidth: "number",
    setFormatPainterModeState: "a user choice",
    setActiveNoteStory: "set when a note story opens",
    setShowPageSetup: "boolean",
    setControlledCommentsRetry: "number",
  },
  "paged-editor/PagedEditor.tsx": {
    setNoteStoryActive: "boolean",
    setLayout:
      "a new layout on every pass; comparing it costs more than a render",
    setBlocks: "set with the layout",
    setMeasures: "set with the layout",
    setShouldCreateHiddenEditorView: "boolean",
    setIsFocused: "boolean",
    setAutocompleteIsStreaming: "boolean",
    setPreviousHfEditMode: "boolean",
  },
  "components/useFolioComments.ts": {
    setShowCommentsSidebar: "boolean",
    setVisibleCommentAuthors: "a user choice",
    setActiveCommentId: "number",
    setInternalComments: "the comment list, set when it changes",
    setIsAddingComment: "boolean",
    setAddCommentYPosition: "number",
    setCommentsLoaded: "boolean",
    setLoadedCommentsNotification: "set once per load",
  },
};

describe("editor host state", () => {
  for (const [file, allowed] of Object.entries(RAW_STATE)) {
    test(`${file} writes object-valued typing state through useEqualState`, () => {
      const source = readFileSync(
        path.join(import.meta.dir, "..", file),
        "utf8",
      );
      const raw = [
        ...source.matchAll(/const \[\w*, (set\w+)\] =\s*useState\b/gu),
      ].map((match) => match[1]);
      expect(raw.filter((setter) => !((setter ?? "") in allowed))).toEqual([]);
      // A listed state that no longer exists leaves the list.
      expect(
        Object.keys(allowed).filter((setter) => !raw.includes(setter)),
      ).toEqual([]);
    });
  }
});
