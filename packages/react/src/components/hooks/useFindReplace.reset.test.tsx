import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef, useImperativeHandle, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import type { Document } from "@stll/folio-core/types/document";
import type { FindMatch } from "../dialogs/findReplaceUtils";
import { findInDocument } from "@stll/folio-core/utils/findReplace";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
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

const replacementFixture = (kind: "field" | "hidden" | "ordinary"): Document => {
  if (kind === "ordinary") return createEmptyDocument({ initialText: "hit" });
  if (kind === "field")
    return {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "complexField",
                  fieldType: "DOCVARIABLE",
                  instruction: "DOCVARIABLE label",
                  fieldCode: [],
                  fieldResult: [{ type: "run", content: [{ type: "text", text: "hit hit" }] }],
                },
                { type: "run", content: [{ type: "text", text: " ordinary hit" }] },
              ],
            },
          ],
        },
      },
    };
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "A" }] },
              {
                type: "deletion",
                info: { id: 1, author: "Author" },
                content: [{ type: "run", content: [{ type: "text", text: "X" }] }],
              },
              { type: "run", content: [{ type: "text", text: "B" }] },
            ],
          },
        ],
      },
    },
  };
};

test.each(["body", "document"] as const)(
  "replacement refuses stale matches before the debounced refresh using %s identity",
  async (source) => {
    let liveDocument = createEmptyDocument({ initialText: "hit" });
    let bodyDoc = toProseDoc(liveDocument);
    const applied: Document[] = [];
    const ref = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: () => (source === "body" ? structuredClone(liveDocument) : liveDocument),
        ...(source === "body" ? { getBodyDoc: () => bodyDoc } : {}),
        handleDocumentChange: (nextDocument) => applied.push(nextDocument),
      });
      useImperativeHandle(ref, () => operations);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      for (const command of ["current", "all"] as const) {
        liveDocument = createEmptyDocument({ initialText: "hit" });
        bodyDoc = toProseDoc(liveDocument);
        await act(async () => {
          expect(operations().handleFind("hit", options)?.totalCount).toBe(1);
        });
        liveDocument = replacementFixture("field");
        bodyDoc = toProseDoc(liveDocument);
        const before = structuredClone(liveDocument);
        await act(async () => {
          if (command === "current") expect(operations().handleReplace("changed")).toBe(false);
          else
            expect(operations().handleReplaceAll("hit", "changed", options)).toEqual({
              replacedCount: 0,
              skippedCount: 1,
            });
        });
        expect(applied).toEqual([]);
        expect(liveDocument).toEqual(before);
        expect(operations().getFindResult()?.totalCount).toBe(3);
        expect(operations().getFindResult()?.matches.at(0)?.replacement).toEqual({
          status: "blocked",
          reason: "field-result",
        });
      }
      liveDocument = createEmptyDocument({ initialText: "hit" });
      bodyDoc = toProseDoc(liveDocument);
      await act(async () => {
        operations().handleFind("hit", options);
      });
      await act(async () => {
        expect(operations().handleReplace("changed")).toBe(true);
      });
      expect(applied).toHaveLength(1);
      expect(findInDocument(applied.at(0), "changed", options)).toHaveLength(1);
    } finally {
      await act(async () => root.unmount());
    }
  },
);

test.each(["current", "all"] as const)(
  "replacement refuses text hidden by a display-mode change before refresh: %s",
  async (command) => {
    const fixture = replacementFixture("hidden");
    const before = structuredClone(fixture);
    const bodyDoc = toProseDoc(fixture);
    const applied: Document[] = [];
    const ref = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = ({ mode }: { mode: "all-markup" | "no-markup" }) => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        documentState: fixture,
        containerRef,
        findReplace,
        displayMode: mode,
        getBodyDoc: () => bodyDoc,
        handleDocumentChange: (nextDocument) => applied.push(nextDocument),
      });
      useImperativeHandle(ref, () => operations);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness mode="all-markup" />));
      await act(async () => {
        expect(operations().handleFind("X", options)?.totalCount).toBe(1);
      });
      await act(async () => root.render(<Harness mode="no-markup" />));
      expect(operations().getFindResult()?.totalCount).toBe(1);
      await act(async () => {
        if (command === "current") expect(operations().handleReplace("changed")).toBe(false);
        else
          expect(operations().handleReplaceAll("X", "changed", options)).toEqual({
            replacedCount: 0,
            skippedCount: 1,
          });
      });
      expect(operations().getFindResult()?.totalCount).toBe(0);
      expect(applied).toEqual([]);
      expect(fixture).toEqual(before);
    } finally {
      await act(async () => root.unmount());
    }
  },
);

test.each(["field", "hidden", "ordinary"] as const)(
  "replacement applies only authored editable ranges: %s",
  async (kind) => {
    const fixture = replacementFixture(kind);
    const before = structuredClone(fixture);
    const ref = createRef<{ operations: ReturnType<typeof useFindReplace>; document: Document }>();
    const applied: Document[] = [];
    const Harness = () => {
      const [currentDocument, setCurrentDocument] = useState(fixture);
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        documentState: currentDocument,
        containerRef,
        findReplace,
        displayMode: "no-markup",
        handleDocumentChange: (nextDocument) => {
          applied.push(nextDocument);
          setCurrentDocument(nextDocument);
        },
      });
      useImperativeHandle(ref, () => ({ operations, document: currentDocument }));
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current?.operations ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        operations().handleFind(kind === "hidden" ? "AB" : "hit", options);
      });
      await act(async () => {
        expect(operations().handleReplace("changed")).toBe(kind === "ordinary");
      });
      if (kind === "ordinary") {
        expect(applied).toHaveLength(1);
        expect(findInDocument(ref.current?.document, "changed", options)).toHaveLength(1);
        expect(findInDocument(ref.current?.document, "hit", options)).toHaveLength(0);
        return;
      }
      expect(applied).toEqual([]);
      await act(async () => {
        expect(
          operations().handleReplaceAll(kind === "hidden" ? "AB" : "hit", "changed", options),
        ).toEqual(
          kind === "field"
            ? { replacedCount: 1, skippedCount: 2 }
            : { replacedCount: 0, skippedCount: 1 },
        );
      });
      if (kind === "hidden") {
        expect(applied).toEqual([]);
        expect(ref.current?.document).toEqual(before);
        return;
      }
      expect(applied).toHaveLength(1);
      const currentParagraph = ref.current?.document.package.document.content.at(0);
      const originalParagraph = before.package.document.content.at(0);
      expect(currentParagraph?.type).toBe("paragraph");
      if (currentParagraph?.type !== "paragraph" || originalParagraph?.type !== "paragraph")
        panic("Replacement fixture paragraph missing");
      expect(currentParagraph.content.at(0)).toEqual(originalParagraph.content.at(0));
      expect(findInDocument(ref.current?.document, "hit", options)).toHaveLength(2);
      expect(findInDocument(ref.current?.document, "changed", options)).toHaveLength(1);
      expect(fixture).toEqual(before);
    } finally {
      await act(async () => root.unmount());
    }
  },
);
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
