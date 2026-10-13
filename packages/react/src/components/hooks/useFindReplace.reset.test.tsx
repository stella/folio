import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef, useImperativeHandle, useRef } from "react";
import { createRoot } from "react-dom/client";
import { EditorState } from "prosemirror-state";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import { schema } from "@stll/folio-core/prosemirror/schema";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import type { Document } from "@stll/folio-core/types/document";
import type { FindMatch } from "../dialogs/findReplaceUtils";
import { useFindReplace as useFindDialogState } from "../dialogs/useFindReplace";
import { useFindReplace } from "./useFindReplace";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const options = { matchCase: false, matchWholeWord: false };
const noop = () => {};
type FindHarnessRef = {
  operations: ReturnType<typeof useFindReplace>;
  matches: FindMatch[];
};

type RecordedReplacement = { matches: readonly FindMatch[]; replacement: string };
const createReplacementRecorder =
  (replacements: RecordedReplacement[]) => (matches: readonly FindMatch[], replacement: string) => {
    replacements.push({ matches, replacement });
    return true;
  };

test("query and document resets clear manager navigation, reactive result, and dialog matches together", async () => {
  const ref = createRef<FindHarnessRef>();
  const fixture = createEmptyDocument({ initialText: "needle needle" });
  const Harness = ({ document }: { document: Document | null }) => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const operations = useFindReplace({
      documentState: document,
      containerRef,
      handleDocumentChange: noop,
      findReplace,
    });
    useImperativeHandle(ref, () => ({ operations, matches: findReplace.state.matches }), [
      operations,
      findReplace.state.matches,
    ]);
    return null;
  };
  const container = document.createElement("div");
  const root = createRoot(container);
  const getOperations = () => ref.current?.operations ?? panic("Find hook handle missing");
  const assertCleared = () => {
    expect(getOperations().currentResult).toBeNull();
    expect(ref.current?.matches).toEqual([]);
    expect(getOperations().handleFindNext()).toBeNull();
    expect(getOperations().handleFindPrevious()).toBeNull();
    expect(getOperations().handleReplace("replacement")).toBe(false);
  };
  try {
    await act(async () => root.render(<Harness document={fixture} />));
    for (const reset of [
      () => getOperations().resetFindResult(),
      () => getOperations().handleFind("", options),
      () => getOperations().handleFind("  ", options),
    ]) {
      await act(async () => {
        getOperations().handleFind("needle", options);
      });
      expect(getOperations().currentResult?.totalCount).toBe(2);
      expect(ref.current?.matches).toHaveLength(2);
      await act(async () => {
        reset();
      });
      assertCleared();
    }
    await act(async () => {
      getOperations().handleFind("needle", options);
    });
    await act(async () => root.render(<Harness document={null} />));
    await act(async () => {
      getOperations().handleFind("needle", options);
    });
    assertCleared();
  } finally {
    await act(async () => root.unmount());
  }
});

test("switching markup views refreshes active results and preserves authored selection ranges", async () => {
  const ref = createRef<FindHarnessRef>();
  const deletion = (schema.marks["deletion"] ?? panic("Missing deletion mark")).create({
    revisionId: 1,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
  });
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("needle", [deletion]), schema.text(" kept")]),
    ]),
  });
  const fixture = createEmptyDocument({ initialText: "needle kept" });
  const selected: FindMatch[] = [];
  const Harness = ({
    markupView,
    replaceEditorMatches,
  }: {
    markupView: DisplayMode;
    replaceEditorMatches?: (matches: readonly FindMatch[], replacement: string) => boolean;
  }) => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const operations = useFindReplace({
      documentState: fixture,
      getEditorState: () => state,
      markupView,
      replaceEditorMatches,
      containerRef,
      handleDocumentChange: noop,
      findReplace,
      selectMatch: (match) => {
        selected.push(match);
        return true;
      },
    });
    useImperativeHandle(ref, () => ({ operations, matches: findReplace.state.matches }));
    return null;
  };
  const root = createRoot(document.createElement("div"));
  const getOperations = () => ref.current?.operations ?? panic("Find hook handle missing");
  try {
    await act(async () => root.render(<Harness markupView="all-markup" />));
    await act(async () => {
      getOperations().handleFind("needle", options);
    });
    expect(getOperations().currentResult?.totalCount).toBe(1);
    await act(async () => root.render(<Harness markupView="no-markup" />));
    expect(getOperations().currentResult?.totalCount).toBe(0);
    expect(getOperations().handleFindNext()).toBeNull();
    await act(async () => root.render(<Harness markupView="original" />));
    expect(getOperations().currentResult?.totalCount).toBe(1);
    expect(selected.at(-1)?.authoredRange).toEqual({ from: 1, to: 7 });
    expect(getOperations().handleReplace("replacement")).toBe(false);
    expect(getOperations().handleReplaceAll("needle", "replacement", options)).toBe(0);
    const replacements: RecordedReplacement[] = [];
    const replaceEditorMatches = createReplacementRecorder(replacements);
    await act(async () =>
      root.render(<Harness markupView="original" replaceEditorMatches={replaceEditorMatches} />),
    );
    await act(async () => {
      expect(getOperations().handleReplace("replacement")).toBe(true);
    });
    expect(replacements.at(-1)).toMatchObject({
      matches: [{ authoredRange: { from: 1, to: 7 } }],
      replacement: "replacement",
    });
    await act(async () => {
      expect(getOperations().handleReplaceAll("needle", "replacement", options)).toBe(1);
    });
    expect(replacements).toHaveLength(2);
  } finally {
    await act(async () => root.unmount());
  }
});
