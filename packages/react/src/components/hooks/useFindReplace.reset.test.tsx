import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef, useImperativeHandle, useRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { IntlProvider } from "use-intl";

import {
  buildRawPackage,
  SECTION,
  WORD_STYLES,
} from "@stll/folio-core/__tests__/documentShapes";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import type { Layout } from "@stll/folio-core/layout-engine/types";
import { DISPLAY_MODES } from "@stll/folio-core/managers/EditorModeManager";
import type {
  DisplayMode,
  EditorReadOnlyReason,
} from "@stll/folio-core/managers/EditorModeManager";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { schema } from "@stll/folio-core/prosemirror/schema";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { findInDocument } from "@stll/folio-core/utils/findReplace";

import {
  buildBoundaryDocument,
  surroundHiddenBoundary,
  ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES,
} from "../../../../../test/findReplacementBoundaryFixtures";
import { FindReplaceDialog } from "../dialogs/FindReplaceDialog";
import type { FindMatch } from "../dialogs/findReplaceUtils";
import { useFindReplace as useFindDialogState } from "../dialogs/useFindReplace";
import { useFindReplace } from "./useFindReplace";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const options = { matchCase: false, matchWholeWord: false };
const noop = () => {};

/** A mounted body editor over `fixture`; the host's document reads it back. */
const mountBody = (fixture: Document) => {
  const hidden = document.createElement("div");
  document.body.append(hidden);
  const view = new EditorView(hidden, {
    state: EditorState.create({ doc: toProseDoc(fixture) }),
  });
  return {
    view,
    read: () =>
      fromProseDoc(view.state.doc, fixture, {
        stylesheetSource: { type: "package" },
      }),
    destroy: () => {
      view.destroy();
      hidden.remove();
    },
  };
};

const pendingOptionCases = [
  {
    option: "matchCase",
    text: "hit HIT",
    query: "hit",
    previousCount: 2,
    refreshedCount: 1,
  },
  {
    option: "matchWholeWord",
    text: "hit hitman",
    query: "hit",
    previousCount: 2,
    refreshedCount: 1,
  },
  {
    option: "useRegex",
    text: "a.b axb",
    query: "a.b",
    previousCount: 1,
    refreshedCount: 2,
  },
] as const;

test.each(
  pendingOptionCases.flatMap((scenario) =>
    (["current", "all"] as const).flatMap((command) =>
      (["unchanged", "changed"] as const).map(
        (replacement) =>
          [scenario.option, command, replacement, scenario] as const,
      ),
    ),
  ),
)(
  "React refuses pending %s changes before %s replacement (%s)",
  async (_option, command, replacement, scenario) => {
    const fixture = createEmptyDocument({ initialText: scenario.text });
    const before = structuredClone(fixture);
    const body = mountBody(fixture);
    const initialDoc = body.view.state.doc;
    const liveOptions = {
      matchCase: false,
      matchWholeWord: false,
      useRegex: false,
    };
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const search = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: body.read,
        getBodyDoc: () => body.view.state.doc,
        getEditorView: () => body.view,
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () =>
      handle.current ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        expect(
          operations().handleFind(scenario.query, liveOptions)?.totalCount,
        ).toBe(scenario.previousCount);
      });
      // A host can mutate options before the dialog's debounced find has run.
      liveOptions[scenario.option] = true;
      const pendingReplacement =
        replacement === "unchanged" ? scenario.query : "changed";
      await act(async () => {
        if (command === "current")
          expect(operations().handleReplace(pendingReplacement)).toBe(false);
        else
          expect(
            operations().handleReplaceAll(
              scenario.query,
              pendingReplacement,
              liveOptions,
            ),
          ).toEqual({
            replacedCount: 0,
            skippedCount: scenario.previousCount,
          });
      });
      expect(body.view.state.doc).toBe(initialDoc);
      expect(fixture).toEqual(before);
      expect(operations().getFindResult()?.totalCount).toBe(
        scenario.refreshedCount,
      );
      operations().handleQueryChange(scenario.query, { ...liveOptions });
      await act(async () => {
        if (command === "current")
          expect(operations().handleReplace("changed")).toBe(true);
        else
          expect(
            operations().handleReplaceAll(scenario.query, "changed", {
              ...liveOptions,
            }),
          ).toEqual({
            replacedCount: scenario.refreshedCount,
            skippedCount: 0,
          });
      });
      expect(body.view.state.doc.textContent).toContain("changed");
    } finally {
      await act(async () => root.unmount());
      body.destroy();
    }
  },
);

test("React uses the live query edit before its next debounced search", async () => {
  const fixture = createEmptyDocument({ initialText: "alpha beta" });
  const body = mountBody(fixture);
  const initialDoc = body.view.state.doc;
  const handle = createRef<ReturnType<typeof useFindReplace>>();
  const Harness = () => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const search = useFindReplace({
      containerRef,
      findReplace,
      getDocumentState: body.read,
      getBodyDoc: () => body.view.state.doc,
      getEditorView: () => body.view,
    });
    useImperativeHandle(handle, () => search);
    return null;
  };
  const root = createRoot(document.createElement("div"));
  const operations = () => handle.current ?? panic("Find hook handle missing");
  try {
    await act(async () => root.render(<Harness />));
    await act(async () => operations().handleFind("alpha", options));
    operations().handleQueryChange("beta", { ...options });
    await act(async () =>
      expect(operations().handleReplace("changed")).toBe(false),
    );
    expect(body.view.state.doc).toBe(initialDoc);
    expect(operations().getFindResult()?.matches.at(0)?.text).toBe("beta");
    await act(async () =>
      expect(operations().handleReplace("changed")).toBe(true),
    );
    expect(body.view.state.doc.textContent).toBe("alpha changed");
  } finally {
    await act(async () => root.unmount());
    body.destroy();
  }
});

test.each(["current", "all"] as const)(
  "React refreshes its result from the edited body after %s replacement",
  async (command) => {
    const body = mountBody(createEmptyDocument({ initialText: "hit" }));
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const search = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: body.read,
        getBodyDoc: () => body.view.state.doc,
        getEditorView: () => body.view,
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () =>
      handle.current ?? panic("Find hook handle missing");
    const replace = (query: string, text: string) =>
      command === "current"
        ? operations().handleReplace(text)
        : operations().handleReplaceAll(query, text, { ...options })
            .replacedCount > 0;
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => operations().handleFind("hit", options));
      await act(async () => expect(replace("hit", "changed")).toBe(true));
      expect(body.view.state.doc.textContent).toBe("changed");
      // The written body answers the active query at once.
      expect(operations().getFindResult()?.totalCount).toBe(0);
      await act(async () => expect(replace("hit", "duplicate")).toBe(false));
      expect(body.view.state.doc.textContent).toBe("changed");

      await act(async () => operations().handleFind("changed", options));
      await act(async () => expect(replace("changed", "committed")).toBe(true));
      expect(body.view.state.doc.textContent).toBe("committed");
    } finally {
      await act(async () => root.unmount());
      body.destroy();
    }
  },
);

test.each(["current", "all"] as const)(
  "React accepts an unchanged %s replacement without an edit",
  async (command) => {
    const fixture = createEmptyDocument({ initialText: "hit" });
    const host = document.createElement("div");
    const hidden = document.createElement("div");
    document.body.append(host, hidden);
    const view = new EditorView(hidden, {
      state: EditorState.create({ doc: toProseDoc(fixture) }),
    });
    const originalState = view.state;
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const search = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: () =>
          fromProseDoc(view.state.doc, fixture, {
            stylesheetSource: { type: "package" },
          }),
        getBodyDoc: () => view.state.doc,
        getEditorView: () => view,
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(host);
    const operations = () =>
      handle.current ?? panic("Find hook handle missing");
    const replace = (text: string) =>
      command === "current"
        ? operations().handleReplace(text)
        : operations().handleReplaceAll("hit", text, { ...options })
            .replacedCount > 0;
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => operations().handleFind("hit", options));
      await act(async () => expect(replace("hit")).toBe(true));
      expect(view.state).toBe(originalState);
      await act(async () => operations().handleFind("hit", { ...options }));
      await act(async () => expect(replace("changed")).toBe(true));
      expect(view.state.doc.textContent).toBe("changed");
    } finally {
      await act(async () => root.unmount());
      view.destroy();
      host.remove();
      hidden.remove();
    }
  },
);

test("React host search changes cannot relabel the mounted dialog result", async () => {
  const fixture = createEmptyDocument({ initialText: "alpha beta" });
  const body = mountBody(fixture);
  const initialDoc = body.view.state.doc;
  const handle = createRef<ReturnType<typeof useFindReplace>>();
  const Harness = () => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const search = useFindReplace({
      containerRef,
      findReplace,
      getDocumentState: body.read,
      getBodyDoc: () => body.view.state.doc,
      getEditorView: () => body.view,
    });
    useImperativeHandle(handle, () => search);
    return (
      <FindReplaceDialog
        isOpen
        replaceMode
        initialSearchText="alpha"
        onClose={noop}
        onFind={search.handleFind}
        onQueryChange={search.handleQueryChange}
        onFindNext={search.handleFindNext}
        onFindPrevious={search.handleFindPrevious}
        onReplace={search.handleReplace}
        onReplaceAll={search.handleReplaceAll}
        currentSearch={search.currentSearch}
      />
    );
  };
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const operations = () => handle.current ?? panic("Find hook handle missing");
  const replaceButton = () =>
    Array.from(host.querySelectorAll<HTMLButtonElement>("button")).find(
      (button) => button.textContent === "Replace current match",
    );
  try {
    await act(async () =>
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <Harness />
        </IntlProvider>,
      ),
    );
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 140));
    });
    expect(replaceButton()?.disabled).toBe(false);
    await act(async () => operations().handleFind("beta", options));
    expect(operations().getFindResult()?.matches.at(0)?.text).toBe("beta");
    expect(host.querySelector("input")?.value).toBe("alpha");
    expect(replaceButton()?.disabled).toBe(true);
    await act(async () => replaceButton()?.click());
    expect(body.view.state.doc).toBe(initialDoc);
  } finally {
    await act(async () => root.unmount());
    host.remove();
    body.destroy();
  }
});

// Hidden-revision fixtures previously contained visible authored text only;
// exercise zero-text carriers through the mounted hook and real PM conversion.
test.each(
  Object.entries(ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES).flatMap(
    ([name, content]) =>
      (["deletion", "moveFrom"] as const).map(
        (revisionType) => [name, revisionType, content] as const,
      ),
  ),
)(
  "React replacement preserves a zero-text protected boundary: %s / %s",
  async (_name, revisionType, content) => {
    const fixture = buildBoundaryDocument(
      surroundHiddenBoundary(content, revisionType),
    );
    const before = structuredClone(fixture);
    const host = document.createElement("div");
    const hidden = document.createElement("div");
    document.body.append(host, hidden);
    const view = new EditorView(hidden, {
      state: EditorState.create({ doc: toProseDoc(fixture) }),
    });
    const originalState = view.state;
    const originalJSON = view.state.doc.toJSON();
    const revision = fixture.package.document.content.at(0);
    if (revision?.type !== "paragraph")
      panic("Boundary fixture paragraph missing");
    const boundary = revision.content.at(1);
    if (boundary?.type !== "deletion" && boundary?.type !== "moveFrom")
      panic("Boundary fixture revision missing");
    let hasRevision = false;
    view.state.doc.descendants((node) => {
      hasRevision ||= node.marks.some(
        (mark) =>
          mark.type.name === "deletion" &&
          mark.attrs["revisionId"] === boundary.info.id,
      );
    });
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const containerRef = useRef<HTMLDivElement>(null);
      const findReplace = useFindDialogState();
      const search = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: () =>
          fromProseDoc(view.state.doc, fixture, {
            stylesheetSource: { type: "package" },
          }),
        getBodyDoc: () => view.state.doc,
        displayMode: "no-markup",
        getEditorView: () => view,
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(host);
    const operations = () =>
      handle.current ?? panic("Find hook handle missing");
    try {
      expect(hasRevision).toBe(true);
      await act(async () => root.render(<Harness />));
      for (const command of ["current", "all"] as const) {
        await act(async () => {
          expect(operations().handleFind("AB", options)?.totalCount).toBe(1);
        });
        expect(
          operations().getFindResult()?.matches.at(0)?.replacement,
        ).toEqual({
          status: "blocked",
          reason: "hidden-content",
        });
        await act(async () => {
          if (command === "current")
            expect(operations().handleReplace("changed")).toBe(false);
          else
            expect(
              operations().handleReplaceAll("AB", "changed", options),
            ).toEqual({
              replacedCount: 0,
              skippedCount: 1,
            });
        });
        expect(view.state).toBe(originalState);
        expect(view.state.doc.toJSON()).toEqual(originalJSON);
        expect(fixture).toEqual(before);
      }
    } finally {
      await act(async () => root.unmount());
      view.destroy();
      host.remove();
      hidden.remove();
    }
  },
);

const replacementFixture = (
  kind: "field" | "hidden" | "ordinary",
): Document => {
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
                  fieldResult: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "hit hit" }],
                    },
                  ],
                },
                {
                  type: "run",
                  content: [{ type: "text", text: " ordinary hit" }],
                },
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
                content: [
                  { type: "run", content: [{ type: "text", text: "X" }] },
                ],
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
    const body = mountBody(createEmptyDocument({ initialText: "hit" }));
    // Another edit replaces the body behind the find result's back.
    const replaceBody = (next: Document) =>
      body.view.dispatch(
        body.view.state.tr.replaceWith(
          0,
          body.view.state.doc.content.size,
          toProseDoc(next).content,
        ),
      );
    const ref = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        containerRef,
        findReplace,
        getDocumentState: body.read,
        ...(source === "body" ? { getBodyDoc: () => body.view.state.doc } : {}),
        getEditorView: () => body.view,
      });
      useImperativeHandle(ref, () => operations);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      for (const command of ["current", "all"] as const) {
        replaceBody(createEmptyDocument({ initialText: "hit" }));
        await act(async () => {
          expect(operations().handleFind("hit", options)?.totalCount).toBe(1);
        });
        replaceBody(replacementFixture("field"));
        const before = body.view.state.doc;
        await act(async () => {
          if (command === "current")
            expect(operations().handleReplace("changed")).toBe(false);
          else
            expect(
              operations().handleReplaceAll("hit", "changed", options),
            ).toEqual({
              replacedCount: 0,
              skippedCount: 1,
            });
        });
        expect(body.view.state.doc).toBe(before);
        expect(operations().getFindResult()?.totalCount).toBe(3);
        expect(
          operations().getFindResult()?.matches.at(0)?.replacement,
        ).toEqual({
          status: "blocked",
          reason: "field-result",
        });
      }
      replaceBody(createEmptyDocument({ initialText: "hit" }));
      await act(async () => {
        operations().handleFind("hit", options);
      });
      await act(async () => {
        expect(operations().handleReplace("changed")).toBe(true);
      });
      expect(body.view.state.doc.textContent).toBe("changed");
    } finally {
      await act(async () => root.unmount());
      body.destroy();
    }
  },
);

test.each(["current", "all"] as const)(
  "replacement refuses text hidden by a display-mode change before refresh: %s",
  async (command) => {
    const fixture = replacementFixture("hidden");
    const before = structuredClone(fixture);
    const bodyDoc = toProseDoc(fixture);
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
      // The display-mode switch refreshes at once, and review views refuse replacement.
      await act(async () => root.render(<Harness mode="no-markup" />));
      expect(operations().getFindResult()?.totalCount).toBe(0);
      await act(async () => {
        if (command === "current")
          expect(operations().handleReplace("changed")).toBe(false);
        else
          expect(
            operations().handleReplaceAll("X", "changed", options),
          ).toEqual({
            replacedCount: 0,
            skippedCount: 0,
          });
      });
      expect(operations().getFindResult()?.totalCount).toBe(0);
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
    const body = mountBody(fixture);
    const initialDoc = body.view.state.doc;
    const ref = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        getDocumentState: body.read,
        getBodyDoc: () => body.view.state.doc,
        getEditorView: () => body.view,
        containerRef,
        findReplace,
        // Review views refuse replacement; only the hidden fixture needs No Markup.
        displayMode: kind === "hidden" ? "no-markup" : "all-markup",
      });
      useImperativeHandle(ref, () => operations);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () => ref.current ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        operations().handleFind(kind === "hidden" ? "AB" : "hit", options);
      });
      await act(async () => {
        expect(operations().handleReplace("changed")).toBe(kind === "ordinary");
      });
      if (kind === "ordinary") {
        expect(body.view.state.doc.textContent).toBe("changed");
        return;
      }
      expect(body.view.state.doc).toBe(initialDoc);
      await act(async () => {
        expect(
          operations().handleReplaceAll(
            kind === "hidden" ? "AB" : "hit",
            "changed",
            options,
          ),
        ).toEqual(
          kind === "field"
            ? { replacedCount: 1, skippedCount: 2 }
            : { replacedCount: 0, skippedCount: 1 },
        );
      });
      if (kind === "hidden") {
        expect(body.view.state.doc).toBe(initialDoc);
        expect(fixture).toEqual(before);
        return;
      }
      // The field keeps its result; only the ordinary text is replaced.
      const written = body.read();
      expect(findInDocument(written, "hit", options)).toHaveLength(2);
      expect(findInDocument(written, "changed", options)).toHaveLength(1);
      const paragraph = written.package.document.content.at(0);
      const originalParagraph = before.package.document.content.at(0);
      if (
        paragraph?.type !== "paragraph" ||
        originalParagraph?.type !== "paragraph"
      )
        panic("Replacement fixture paragraph missing");
      expect(paragraph.content.at(0)?.type).toBe(
        originalParagraph.content.at(0)?.type,
      );
      expect(fixture).toEqual(before);
    } finally {
      await act(async () => root.unmount());
      body.destroy();
    }
  },
);
type FindHarnessRef = {
  operations: ReturnType<typeof useFindReplace>;
  matches: FindMatch[];
};

test.each(DISPLAY_MODES.filter((mode) => mode !== "all-markup"))(
  "initial %s searches defer until the lazy body editor has a state",
  async (markupView) => {
    const revisionAttrs = {
      revisionId: 1,
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
    };
    const insertion = (
      schema.marks["insertion"] ?? panic("Missing insertion mark")
    ).create(revisionAttrs);
    const deletion = (
      schema.marks["deletion"] ?? panic("Missing deletion mark")
    ).create(revisionAttrs);
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("inserted", [insertion]),
          schema.text(" deleted", [deletion]),
          schema.text(" kept"),
        ]),
      ]),
    });
    const fixture = createEmptyDocument({
      initialText: "inserted deleted kept",
    });
    const hiddenQueries = {
      "all-markup": null,
      "simple-markup": "deleted",
      "no-markup": "deleted",
      original: "inserted",
    } as const satisfies Record<DisplayMode, string | null>;
    const hiddenQuery = hiddenQueries[markupView];
    let editorState: EditorState | null = null;
    const ref = createRef<FindHarnessRef>();
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        documentState: fixture,
        getEditorState: () => editorState,
        displayMode: markupView,
        containerRef,
        findReplace,
      });
      useImperativeHandle(ref, () => ({
        operations,
        matches: findReplace.state.matches,
      }));
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const getOperations = () =>
      ref.current?.operations ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        expect(getOperations().handleFind(hiddenQuery, options)).toBeNull();
      });
      expect(ref.current?.matches).toEqual([]);
      expect(getOperations().handleFindNext()).toBeNull();
      editorState = state;
      await act(async () => {
        expect(
          getOperations().handleFind(hiddenQuery, options)?.totalCount,
        ).toBe(0);
        expect(getOperations().handleFind("kept", options)?.totalCount).toBe(1);
      });
    } finally {
      await act(async () => root.unmount());
    }
  },
);

test("query and document resets clear manager navigation, reactive result, and dialog matches together", async () => {
  const ref = createRef<FindHarnessRef>();
  const fixture = createEmptyDocument({ initialText: "needle needle" });
  const Harness = ({ document }: { document: Document | null }) => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const operations = useFindReplace({
      documentState: document,
      containerRef,
      findReplace,
    });
    useImperativeHandle(
      ref,
      () => ({ operations, matches: findReplace.state.matches }),
      [operations, findReplace.state.matches],
    );
    return null;
  };
  const container = document.createElement("div");
  const root = createRoot(container);
  const getOperations = () =>
    ref.current?.operations ?? panic("Find hook handle missing");
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
            content: [
              { type: "run", content: [{ type: "text", text: "needle" }] },
            ],
          },
        ],
      },
    ];
    const before = structuredClone(fixture);
    const ref = createRef<FindHarnessRef>();
    const notifications: ReturnType<typeof useFindReplace>["currentResult"][] =
      [];
    const Harness = () => {
      const findReplace = useFindDialogState();
      const containerRef = useRef<HTMLDivElement>(null);
      const operations = useFindReplace({
        documentState: fixture,
        containerRef,
        findReplace,
        readOnlyReason: readOnly ? "host-read-only" : null,
        onFindResultChange: (result) => notifications.push(result),
      });
      useImperativeHandle(ref, () => ({
        operations,
        matches: findReplace.state.matches,
      }));
      return null;
    };
    const root = createRoot(document.createElement("div"));
    const operations = () =>
      ref.current?.operations ?? panic("Find hook handle missing");
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        expect(operations().handleFind("needle", options)?.totalCount).toBe(3);
      });
      await act(async () => {
        operations().handleFindNext();
      });
      await act(async () => {
        expect(operations().handleFindNext()?.story).toEqual({
          kind: "footnote",
          id: 1,
        });
      });
      expect(operations().getFindResult()?.currentIndex).toBe(2);
      expect(operations().handleReplace("changed")).toBe(false);
      await act(async () => {
        operations().handleFindNext();
      });
      await act(async () => {
        operations().handleFindPrevious();
      });
      expect(notifications.map((result) => result?.currentIndex)).toEqual([
        0, 1, 2, 0, 2,
      ]);
      await act(async () => {
        operations().handleFind("absent", options);
      });
      expect(operations().getFindResult()).toMatchObject({
        totalCount: 0,
        currentIndex: 0,
      });
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

// Adapter navigation previously covered only already-mounted story text. Model
// a virtual page shell and collapsed reply with production DOM attributes.
const navigationFixture = (kind: "body" | "header" | "reply"): Document => {
  const doc = createEmptyDocument({
    initialText: kind === "body" ? "needle" : "body",
  });
  if (kind === "header") {
    doc.package.headers = new Map([
      [
        "h-default",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "needle" }] },
              ],
            },
          ],
        },
      ],
    ]);
    doc.package.document.sections = [
      {
        properties: {
          headerReferences: [{ type: "default", rId: "h-default" }],
        },
        content: [],
      },
    ];
  }
  if (kind === "reply")
    doc.package.document.comments = [
      {
        id: 7,
        author: "Author",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "root" }] },
            ],
          },
        ],
      },
      {
        id: 8,
        parentId: 7,
        author: "Author",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "needle" }] },
            ],
          },
        ],
      },
    ];
  return doc;
};

const navigationLayout = (kind: "body" | "header"): Layout => ({
  pageSize: { w: 800, h: 1000 },
  pages: [1, 2].map((number) => ({
    number,
    logicalNumber: number,
    sectionPageNumber: number,
    size: { w: 800, h: 1000 },
    margins: { top: 20, right: 20, bottom: 20, left: 20 },
    fragments:
      number === 2 && kind === "body"
        ? [
            {
              kind: "paragraph",
              blockId: "block-1",
              x: 20,
              y: 20,
              width: 760,
              height: 18,
              fromLine: 0,
              toLine: 1,
              pmStart: 0,
              pmEnd: 8,
            },
          ]
        : [],
    headerFooterRefs:
      number === 2 && kind === "header" ? { headerDefault: "h-default" } : {},
  })),
});

const navigationViewport = () => {
  const container = document.createElement("div");
  container.dataset["folioScroll"] = "";
  Object.defineProperties(container, {
    clientHeight: { value: 100 },
    clientWidth: { value: 500 },
  });
  container.getBoundingClientRect = () => new DOMRect(0, 0, 500, 100);
  const shell = document.createElement("div");
  shell.className = "layout-page";
  shell.dataset["pageNumber"] = "2";
  shell.getBoundingClientRect = () => new DOMRect(0, 1000, 800, 1000);
  container.append(shell);
  document.body.append(container);
  return { container, shell };
};

const mountNavigationMatch = (shell: HTMLElement, kind: "body" | "header") => {
  const span = document.createElement("span");
  span.className = "layout-run-tab";
  span.dataset["pmStart"] = "1";
  span.dataset["pmEnd"] = "7";
  if (kind === "header") {
    span.dataset["story"] = "header";
    span.dataset["hfRid"] = "h-default";
  }
  span.getBoundingClientRect = () => new DOMRect(30, 1020, 60, 18);
  shell.append(span);
};

test.each(["body", "header"] as const)(
  "React navigation mounts an unmounted page for a %s match",
  async (kind) => {
    const fixture = navigationFixture(kind);
    const { container, shell } = navigationViewport();
    const layout = navigationLayout(kind);
    const bodyDoc = toProseDoc(fixture);
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const scroll = spyOn(container, "scrollTo").mockImplementation(() => {
      if (!shell.querySelector("span"))
        queueMicrotask(() => mountNavigationMatch(shell, kind));
    });
    const Harness = () => {
      const containerRef = useRef(container);
      const findReplace = useFindDialogState();
      const search = useFindReplace({
        containerRef,
        findReplace,
        documentState: fixture,
        getBodyDoc: () => bodyDoc,
        getLayout: () => layout,
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        expect(handle.current?.handleFind("needle", options)?.totalCount).toBe(
          1,
        );
        await new Promise((resolve) => setTimeout(resolve, 40));
      });
      expect(scroll).toHaveBeenCalled();
      expect(shell.querySelector("[data-pm-start]")).not.toBeNull();
      expect(
        container.querySelector("[data-find-match-index='0']"),
      ).not.toBeNull();
      expect(scroll.mock.calls.length).toBeGreaterThan(1);
    } finally {
      await act(async () => root.unmount());
      scroll.mockRestore();
      container.remove();
    }
  },
);

test.each(["collapsed", "visible", "deferred"] as const)(
  "React activates a %s reply before revealing its geometry",
  async (visibility) => {
    const fixture = navigationFixture("reply");
    const bodyDoc = toProseDoc(fixture);
    const { container } = navigationViewport();
    const reply = document.createElement("p");
    reply.dataset["findCommentId"] = "8";
    reply.dataset["findCommentParagraph"] = "0";
    reply.textContent = "needle";
    container.append(reply);
    const events: string[] = [];
    let active = false;
    let committed = visibility === "visible";
    const createRange = document.createRange.bind(document);
    const rangeFactory = spyOn(document, "createRange").mockImplementation(
      () => {
        const range = createRange();
        Object.defineProperty(range, "getClientRects", {
          value: () => {
            events.push(active ? "measure-active" : "measure-inactive");
            return committed ? [new DOMRect(30, 200, 60, 18)] : [];
          },
        });
        return range;
      },
    );
    const scroll = spyOn(container, "scrollTo").mockImplementation(() => {
      events.push(active ? "scroll-active" : "scroll-inactive");
    });
    const handle = createRef<ReturnType<typeof useFindReplace>>();
    const Harness = () => {
      const containerRef = useRef(container);
      const findReplace = useFindDialogState();
      const search = useFindReplace({
        containerRef,
        findReplace,
        documentState: fixture,
        getBodyDoc: () => bodyDoc,
        selectMatch: (match) => {
          expect(match.commentId).toBe(8);
          events.push("activate");
          active = true;
          const commit = () => {
            committed = true;
            events.push("commit");
            reply.dataset["replyExpanded"] = "true";
          };
          if (visibility === "deferred") queueMicrotask(commit);
          else commit();
          return true;
        },
      });
      useImperativeHandle(handle, () => search);
      return null;
    };
    const root = createRoot(document.createElement("div"));
    try {
      await act(async () => root.render(<Harness />));
      await act(async () => {
        handle.current?.handleFind("needle", options);
        if (visibility === "deferred") {
          expect(events).toContain("activate");
          expect(events).not.toContain("commit");
          expect(scroll).not.toHaveBeenCalled();
        }
        await new Promise((resolve) => setTimeout(resolve, 40));
      });
      expect(events).toContain("activate");
      expect(events).not.toContain("scroll-inactive");
      expect(
        container.querySelector(
          "[data-folio-find-comment-overlay] [data-find-match-index='0']",
        ),
      ).not.toBeNull();
      expect(events).toContain("scroll-active");
      expect(events.indexOf("scroll-active")).toBeGreaterThan(
        events.indexOf("commit"),
      );
    } finally {
      await act(async () => root.unmount());
      rangeFactory.mockRestore();
      scroll.mockRestore();
      container.remove();
    }
  },
);

test("React keeps the current match after an older unmounted page finishes mounting", async () => {
  const fixture = createEmptyDocument({ initialText: "needle needle" });
  const bodyDoc = toProseDoc(fixture);
  const { container, shell } = navigationViewport();
  const mountedPage = document.createElement("div");
  mountedPage.className = "layout-page";
  const currentSpan = document.createElement("span");
  currentSpan.className = "layout-run-tab";
  currentSpan.dataset["pmStart"] = "8";
  currentSpan.dataset["pmEnd"] = "14";
  currentSpan.getBoundingClientRect = () => new DOMRect(30, 200, 60, 18);
  mountedPage.append(currentSpan);
  container.append(mountedPage);
  const scroll = spyOn(container, "scrollTo").mockImplementation(() => {});
  const handle = createRef<ReturnType<typeof useFindReplace>>();
  const layout = navigationLayout("body");
  const Harness = () => {
    const containerRef = useRef(container);
    const findReplace = useFindDialogState();
    const search = useFindReplace({
      containerRef,
      findReplace,
      documentState: fixture,
      getBodyDoc: () => bodyDoc,
      getLayout: () => layout,
    });
    useImperativeHandle(handle, () => search);
    return null;
  };
  const root = createRoot(document.createElement("div"));
  try {
    await act(async () => root.render(<Harness />));
    await act(async () => {
      expect(handle.current?.handleFind("needle", options)?.totalCount).toBe(2);
      expect(handle.current?.handleFindNext()?.startOffset).toBe(7);
    });
    const completedScrolls = scroll.mock.calls.length;
    mountNavigationMatch(shell, "body");
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 40));
    });
    expect(scroll.mock.calls.length).toBe(completedScrolls);
    expect(handle.current?.getFindResult()?.currentIndex).toBe(1);
    expect(
      container
        .querySelector("[data-find-match-index='0']")
        ?.getAttribute("data-current"),
    ).toBe("false");
    expect(
      container
        .querySelector("[data-find-match-index='1']")
        ?.getAttribute("data-current"),
    ).toBe("true");
  } finally {
    await act(async () => root.unmount());
    scroll.mockRestore();
    container.remove();
  }
});

test("switching markup views refreshes active results and preserves authored selection ranges", async () => {
  const ref = createRef<FindHarnessRef>();
  const deletion = (
    schema.marks["deletion"] ?? panic("Missing deletion mark")
  ).create({
    revisionId: 1,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
  });
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("needle", [deletion]),
        schema.text(" kept"),
      ]),
    ]),
  });
  const fixture = createEmptyDocument({ initialText: "needle kept" });
  const Harness = ({ markupView }: { markupView: DisplayMode }) => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const operations = useFindReplace({
      documentState: fixture,
      getEditorState: () => state,
      displayMode: markupView,
      containerRef,
      findReplace,
    });
    useImperativeHandle(ref, () => ({
      operations,
      matches: findReplace.state.matches,
    }));
    return null;
  };
  const root = createRoot(document.createElement("div"));
  const getOperations = () =>
    ref.current?.operations ?? panic("Find hook handle missing");
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
    expect(getOperations().currentResult?.matches.at(0)?.authoredRange).toEqual(
      { from: 1, to: 7 },
    );
    expect(getOperations().handleReplace("replacement")).toBe(false);
    expect(
      getOperations().handleReplaceAll("needle", "replacement", options),
    ).toEqual({
      replacedCount: 0,
      skippedCount: 1,
    });
  } finally {
    await act(async () => root.unmount());
  }
});

test("replacement spanning a hidden revision is refused without changing saved bytes", async () => {
  const input = await buildRawPackage({
    styles: WORD_STYLES,
    body:
      '<w:p><w:r><w:t>left</w:t></w:r><w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> hidden </w:delText></w:r></w:del><w:r><w:t>right</w:t></w:r></w:p>' +
      SECTION(),
  });
  const parsed = await parseDocx(input, {
    preloadFonts: false,
    detectVariables: false,
  });
  const body = mountBody(parsed);
  const authored = body.view.state.doc;
  const savedBefore = new Uint8Array(
    await repackDocx(body.read(), { updateModifiedDate: false }),
  );
  const refusals: EditorReadOnlyReason[] = [];
  const ref = createRef<FindHarnessRef>();
  const Harness = ({ markupView }: { markupView: DisplayMode }) => {
    const findReplace = useFindDialogState();
    const containerRef = useRef<HTMLDivElement>(null);
    const operations = useFindReplace({
      getDocumentState: body.read,
      getEditorState: () => body.view.state,
      getEditorView: () => body.view,
      displayMode: markupView,
      onReadonlyEditAttempt: (reason) => refusals.push(reason),
      containerRef,
      findReplace,
    });
    useImperativeHandle(ref, () => ({
      operations,
      matches: findReplace.state.matches,
    }));
    return null;
  };
  const root = createRoot(document.createElement("div"));
  const getOperations = () =>
    ref.current?.operations ?? panic("Find hook handle missing");
  try {
    await act(async () => root.render(<Harness markupView="no-markup" />));
    await act(async () => {
      expect(getOperations().handleFind("leftright", options)?.totalCount).toBe(
        1,
      );
    });
    expect(ref.current?.matches.at(0)?.authoredRange).toEqual({
      from: 1,
      to: 18,
    });
    await act(async () => {
      expect(getOperations().handleReplace("replacement")).toBe(false);
      expect(
        getOperations().handleReplaceAll("leftright", "replacement", options),
      ).toEqual({
        replacedCount: 0,
        skippedCount: 1,
      });
    });
    expect(body.view.state.doc).toBe(authored);
    expect(
      new Uint8Array(
        await repackDocx(body.read(), { updateModifiedDate: false }),
      ),
    ).toEqual(savedBefore);
    expect(refusals).toEqual(["review-display-mode", "review-display-mode"]);
    await act(async () => root.render(<Harness markupView="all-markup" />));
    await act(async () => {
      getOperations().handleFind("left", options);
      expect(getOperations().handleReplace("replacement")).toBe(true);
    });
    expect(body.view.state.doc.textContent).toBe("replacement hidden right");
    await act(async () => {
      getOperations().handleFind("right", options);
      expect(getOperations().handleReplaceAll("right", "end", options)).toEqual(
        {
          replacedCount: 1,
          skippedCount: 0,
        },
      );
    });
    expect(body.view.state.doc.textContent).toBe("replacement hidden end");
  } finally {
    await act(async () => root.unmount());
    body.destroy();
  }
});
