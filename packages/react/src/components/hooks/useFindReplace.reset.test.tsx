import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef, useImperativeHandle, useRef } from "react";
import { createRoot } from "react-dom/client";
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

// Host search and dialog navigation share the manager in both editor modes.
test.each([true, false])(
  "host counts and story navigation work with readOnly=%s",
  async (readOnly) => {
    const fixture = createEmptyDocument({ initialText: "needle needle" });
    fixture.package.footnotes = [
      {
        type: "footnote",
        id: 1,
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "needle" }] }],
          },
        ],
      },
    ];
    const before = structuredClone(fixture);
    const ref = createRef<FindHarnessRef>();
    const notifications: ReturnType<typeof useFindReplace>["currentResult"][] = [];
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        documentState: fixture,
        containerRef,
        handleDocumentChange: noop,
        findReplace,
        readOnly,
        onFindResultChange: (result) => notifications.push(result),
      });
      useImperativeHandle(ref, () => ({ operations, matches: findReplace.state.matches }));
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current?.operations ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        expect(operations().handleFind("needle", options)?.totalCount).toBe(3);
      });
      await act(async () => {
        operations().handleFindNext();
      });
      await act(async () => {
        expect(operations().handleFindNext()?.story).toEqual({ kind: "footnote", id: 1 });
      });
      expect(operations().getFindResult()?.currentIndex).toBe(2);
      expect(operations().handleReplace("changed")).toBe(false);
      await act(async () => {
        operations().handleFindNext();
      });
      await act(async () => {
        operations().handleFindPrevious();
      });
      expect(notifications.map((result) => result?.currentIndex)).toEqual([0, 1, 2, 0, 2]);
      await act(async () => {
        operations().handleFind("absent", options);
      });
      expect(operations().getFindResult()).toMatchObject({ totalCount: 0, currentIndex: 0 });
      await act(async () => {
        operations().resetFindResult();
      });
      expect(notifications.at(-1)).toBeNull();
      expect(fixture).toEqual(before);
    } finally {
      await act(async () => root.unmount());
    }
  },
);
