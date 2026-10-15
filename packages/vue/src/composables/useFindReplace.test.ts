import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, spyOn, test } from "bun:test";
import { panic } from "better-result";
import type { Layout } from "@stll/folio-core/layout-engine/types";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { DISPLAY_MODES } from "@stll/folio-core/managers/EditorModeManager";
import type {
  DisplayMode,
  EditorReadOnlyReason,
} from "@stll/folio-core/managers/EditorModeManager";
import { schema } from "@stll/folio-core/prosemirror/schema";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { buildRawPackage, SECTION, WORD_STYLES } from "@stll/folio-core/__tests__/documentShapes";
import type {
  ComplexField,
  ParagraphContent,
  Run,
  SimpleField,
} from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import type { FindOptions, FindResult } from "@stll/folio-core/utils/findReplace";
import type { UseFindReplaceReturn } from "./useFindReplace";
import { getFindParagraphText } from "@stll/folio-core/utils/findReplace";
import {
  buildBoundaryDocument,
  surroundHiddenBoundary,
  ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES,
} from "../../../../test/findReplacementBoundaryFixtures";

const { createApp, defineComponent, effectScope, h, nextTick, ref, shallowRef } =
  await import("vue");
const { useFindReplace } = await import("./useFindReplace");

afterAll(() => GlobalRegistrator.unregister());

// The previous adapter searched only the body and selected the result in PM.
// Exercise host notifications and wraparound with a secondary story while
// asserting that search leaves both document content and editor selection intact.
test.each([true, false])("Vue search is view-only with readOnly=%s", (readOnly) => {
  const doc = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "hit HIT" }] }],
          },
        ],
      },
      footnotes: [
        {
          type: "footnote",
          id: 1,
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "hit" }] }],
            },
          ],
        },
      ],
    },
  } satisfies Document;
  const before = structuredClone(doc);
  const host = document.createElement("div");
  const hidden = document.createElement("div");
  document.body.append(host, hidden);
  const view = new EditorView(hidden, { state: EditorState.create({ doc: toProseDoc(doc) }) });
  const state = view.state;
  const notifications: (FindResult | null)[] = [];
  const holder: { search: UseFindReplaceReturn | null } = { search: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.search = useFindReplace({
          editorView: shallowRef(view),
          pagesRef: shallowRef(null),
          layout: shallowRef(null),
          getDocument: () => doc,
          displayMode: ref("no-markup"),
          readOnly: ref(readOnly),
          onResultChange: (result) => notifications.push(result),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(host);
  try {
    const search = holder.search ?? panic("Search composable was not mounted");
    expect(search.find("hit")).toMatchObject({ totalCount: 3, currentIndex: 0 });
    expect(search.findNext()?.story).toBe("main");
    expect(search.findNext()?.story).toEqual({ kind: "footnote", id: 1 });
    expect(search.getFindResult()?.currentIndex).toBe(2);
    expect(search.findNext()?.story).toBe("main");
    expect(search.findPrevious()?.story).toEqual({ kind: "footnote", id: 1 });
    search.replaceText.value = "replacement";
    expect(search.replaceCurrent()).toBe(false);
    expect(search.find("hit", { matchCase: true })?.totalCount).toBe(2);
    expect(view.state).toBe(state);
    expect(doc).toEqual(before);
    search.clear();
    expect(search.getFindResult()).toBeNull();
    expect(notifications.at(-1)).toBeNull();
    expect(notifications.slice(0, 5).map((result) => result?.currentIndex)).toEqual([
      0, 1, 2, 0, 2,
    ]);
  } finally {
    app.unmount();
    view.destroy();
    host.remove();
    hidden.remove();
  }
});

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

// Review views refuse replacement; hidden-content fixtures search No Markup.
const withReplacementEditor = (
  contentOrDocument: ParagraphContent[] | Document,
  check: (search: UseFindReplaceReturn, view: EditorView) => void,
  displayMode: DisplayMode = "all-markup",
) => {
  const base = Array.isArray(contentOrDocument)
    ? ({
        package: { document: { content: [{ type: "paragraph", content: contentOrDocument }] } },
      } satisfies Document)
    : contentOrDocument;
  const host = document.createElement("div");
  const hidden = document.createElement("div");
  document.body.append(host, hidden);
  const view = new EditorView(hidden, { state: EditorState.create({ doc: toProseDoc(base) }) });
  const holder: { search: UseFindReplaceReturn | null } = { search: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.search = useFindReplace({
          editorView: shallowRef(view),
          pagesRef: shallowRef(null),
          layout: shallowRef(null),
          getDocument: () =>
            fromProseDoc(view.state.doc, base, { stylesheetSource: { type: "package" } }),
          displayMode: ref(displayMode),
          readOnly: ref(false),
          onResultChange: () => {},
        });
        return () => h("div");
      },
    }),
  );
  app.mount(host);
  try {
    check(holder.search ?? panic("Search composable was not mounted"), view);
  } finally {
    app.unmount();
    view.destroy();
    host.remove();
    hidden.remove();
  }
};

const cachedFields = {
  simple: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF target",
    content: [textRun("42")],
  },
  complex: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF target",
    fieldCode: [],
    fieldResult: [textRun("42")],
  },
} satisfies Record<string, SimpleField | ComplexField>;

// Search fixtures previously checked projection and offsets without executing
// adapter transactions. Exercise both field forms and mixed eligible result
// sets against real PM state so a field atom can never be deleted by replace.
test.each(Object.entries(cachedFields))(
  "Vue refuses replacing a %s field substring",
  (_name, field) => {
    withReplacementEditor([field], (search, view) => {
      const before = view.state;
      expect(search.find("4")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "field-result",
      });
      expect(search.canReplaceCurrent.value).toBe(false);
      search.replaceText.value = "9";
      expect(search.replaceCurrent()).toBe(false);
      expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(view.state).toBe(before);
    });
  },
);

test.each(["fields-only", "mixed"] as const)(
  "Vue replace all reports skipped fields in a %s result set",
  (kind) => {
    const content: ParagraphContent[] = [cachedFields.simple, cachedFields.complex];
    if (kind === "mixed") content.push(textRun(" 4"));
    withReplacementEditor(content, (search, view) => {
      const fieldsBefore: unknown[] = [];
      view.state.doc.descendants((node) => {
        if (node.type.name === "field") fieldsBefore.push(node.toJSON());
      });
      expect(fieldsBefore).toHaveLength(2);
      const before = view.state;
      expect(search.find("4")?.totalCount).toBe(kind === "mixed" ? 3 : 2);
      search.replaceText.value = "9";
      expect(search.replaceAll()).toEqual({
        replacedCount: kind === "mixed" ? 1 : 0,
        skippedCount: 2,
      });
      const fieldsAfter: unknown[] = [];
      view.state.doc.descendants((node) => {
        if (node.type.name === "field") fieldsAfter.push(node.toJSON());
      });
      expect(fieldsAfter).toEqual(fieldsBefore);
      if (kind === "fields-only") expect(view.state).toBe(before);
      else expect(view.state.doc.textContent).toBe("4242 9");
    });
  },
);

test("Vue refuses a replacement crossing hidden deleted content", () => {
  withReplacementEditor(
    [
      textRun("A"),
      { type: "deletion", info: { id: 1, author: "Author" }, content: [textRun("X")] },
      textRun("B"),
    ],
    (search, view) => {
      const before = view.state;
      expect(search.find("AB")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "hidden-content",
      });
      expect(search.canReplaceCurrent.value).toBe(false);
      search.replaceText.value = "C";
      expect(search.replaceCurrent()).toBe(false);
      expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(view.state).toBe(before);
      expect(view.state.doc.textContent).toBe("AXB");
    },
    "no-markup",
  );
});

test("Vue replaces an ordinary editable match", () => {
  withReplacementEditor([textRun("before hit after")], (search, view) => {
    expect(search.find("hit")?.totalCount).toBe(1);
    expect(search.currentMatch.value?.replacement).toEqual({ status: "allowed" });
    expect(search.canReplaceCurrent.value).toBe(true);
    search.replaceText.value = "changed";
    expect(search.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("before changed after");
    expect(search.getFindResult()?.totalCount).toBe(0);
  });
});

test.each(["current", "all"] as const)(
  "Vue accepts %s exact-text replacement without dispatch and permits a later edit",
  (operation) => {
    withReplacementEditor([textRun("hit hit")], (search, view) => {
      expect(search.find("hit")?.totalCount).toBe(2);
      const before = view.state;
      const beforeJson = before.doc.toJSON();
      search.replaceText.value = "hit";
      if (operation === "current") expect(search.replaceCurrent()).toBe(true);
      else expect(search.replaceAll()).toEqual({ replacedCount: 2, skippedCount: 0 });
      expect(view.state).toBe(before);
      expect(view.state.doc.toJSON()).toEqual(beforeJson);
      search.replaceText.value = "changed";
      if (operation === "current") {
        expect(search.replaceCurrent()).toBe(true);
        expect(view.state.doc.textContent).toBe("changed hit");
      } else {
        expect(search.replaceAll()).toEqual({ replacedCount: 2, skippedCount: 0 });
        expect(view.state.doc.textContent).toBe("changed changed");
      }
    });
  },
);

test("Vue mixed replace all leaves an exact-text source run untouched", () => {
  withReplacementEditor(
    [
      { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "hit" }] },
      { type: "run", formatting: { italic: true }, content: [{ type: "text", text: " Hit" }] },
    ],
    (search, view) => {
      expect(search.find("hit")?.totalCount).toBe(2);
      const firstRun = view.state.doc.firstChild?.firstChild;
      expect(firstRun).toBeDefined();
      search.replaceText.value = "hit";
      expect(search.replaceAll()).toEqual({ replacedCount: 2, skippedCount: 0 });
      expect(view.state.doc.textContent).toBe("hit hit");
      expect(view.state.doc.firstChild?.firstChild).toBe(firstRun);
    },
  );
});

for (const [kind, content] of Object.entries(ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES)) {
  for (const revisionType of ["deletion", "moveFrom"] as const) {
    test.each(["current", "all"] as const)(
      `Vue refuses %s replacement across hidden zero-text ${kind} in ${revisionType} without losing source metadata`,
      (operation) => {
        expect(getFindParagraphText({ type: "paragraph", content })).toBe("");
        const fixture = buildBoundaryDocument(surroundHiddenBoundary(content, revisionType));
        withReplacementEditor(
          fixture,
          (search, view) => {
            const before = view.state;
            const beforeJson = before.doc.toJSON();
            expect(JSON.stringify(beforeJson)).toContain("1701037");
            expect(search.find("AB")?.totalCount).toBe(1);
            search.replaceText.value = "C";
            if (operation === "current") expect(search.replaceCurrent()).toBe(false);
            else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
            expect(view.state).toBe(before);
            expect(view.state.doc.toJSON()).toEqual(beforeJson);
            expect(search.currentMatch.value?.replacement).toEqual({
              status: "blocked",
              reason: "hidden-content",
            });
            expect(search.canReplaceCurrent.value).toBe(false);
          },
          "no-markup",
        );
      },
    );
  }
}

// A visible atom with no search text still sits between the matched letters in
// PM; replacing the enclosing range would delete it.
for (const [kind, content] of Object.entries(ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES)) {
  test.each(["current", "all"] as const)(
    `Vue %s replacement keeps a visible zero-text ${kind} between the matched text`,
    (operation) => {
      const fixture = buildBoundaryDocument([textRun("A"), ...content, textRun("B")]);
      withReplacementEditor(fixture, (search, view) => {
        const before = view.state;
        const beforeJson = before.doc.toJSON();
        expect(search.find("AB")?.totalCount).toBe(1);
        expect(search.currentMatch.value?.replacement).toEqual({
          status: "blocked",
          reason:
            kind === "simpleField" || kind === "complexField" ? "field-result" : "atomic-content",
        });
        expect(search.canReplaceCurrent.value).toBe(false);
        search.replaceText.value = "C";
        if (operation === "current") expect(search.replaceCurrent()).toBe(false);
        else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
        expect(view.state).toBe(before);
        expect(view.state.doc.toJSON()).toEqual(beforeJson);
      });
    },
  );
}

// A debounced refresh can be pending when a host changes the current PM doc.
// Re-check the result metadata before any transaction crosses that boundary.
test.each(["current", "all"] as const)(
  "Vue refreshes eligibility before %s replacement after an edit",
  (command) => {
    withReplacementEditor([textRun("42")], (search, view) => {
      expect(search.find("4")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({ status: "allowed" });
      const fieldDoc = toProseDoc({
        package: { document: { content: [{ type: "paragraph", content: [cachedFields.simple] }] } },
      });
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, fieldDoc.content));
      const before = view.state;
      search.replaceText.value = "9";
      if (command === "current") expect(search.replaceCurrent()).toBe(false);
      else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "field-result",
      });
      expect(view.state).toBe(before);
    });
  },
);

type MatchingOptionFixture = {
  text: string;
  query: string;
  initialOptions: Partial<FindOptions>;
  change: (search: UseFindReplaceReturn) => void;
};

const matchingOptionFixtures = {
  matchCase: {
    text: "Hit",
    query: "hit",
    initialOptions: {},
    change: (search) => {
      search.options.matchCase = true;
    },
  },
  matchWholeWord: {
    text: "hitched",
    query: "hit",
    initialOptions: {},
    change: (search) => {
      search.options.matchWholeWord = true;
    },
  },
  useRegex: {
    text: "hit",
    query: "^hit",
    initialOptions: { useRegex: true },
    change: (search) => {
      search.options.useRegex = false;
    },
  },
} satisfies Record<
  keyof Pick<FindOptions, "matchCase" | "matchWholeWord" | "useRegex">,
  MatchingOptionFixture
>;

for (const [option, fixture] of Object.entries(matchingOptionFixtures)) {
  test.each(["current", "all"] as const)(
    `Vue refreshes a changed ${option} before %s replacement while search is debounced`,
    (operation) => {
      withReplacementEditor([textRun(fixture.text)], (search, view) => {
        expect(search.find(fixture.query, fixture.initialOptions)?.totalCount).toBe(1);
        expect(search.canReplaceCurrent.value).toBe(true);
        const before = view.state;
        const beforeJson = before.doc.toJSON();
        fixture.change(search);
        expect(search.canReplaceCurrent.value).toBe(false);
        search.scheduleSearch();
        search.replaceText.value = "replacement";
        if (operation === "current") expect(search.replaceCurrent()).toBe(false);
        else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
        expect(view.state).toBe(before);
        expect(view.state.doc.toJSON()).toEqual(beforeJson);
        expect(search.getFindResult()?.totalCount).toBe(0);
      });
    },
  );
}

// Adapter navigation previously covered only already-mounted story text. Model
// a virtual page shell and collapsed reply with production DOM attributes.
const navigationFixture = (kind: "body" | "header" | "reply"): Document => {
  const doc = createEmptyDocument({ initialText: kind === "body" ? "needle" : "body" });
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
              content: [{ type: "run", content: [{ type: "text", text: "needle" }] }],
            },
          ],
        },
      ],
    ]);
    doc.package.document.sections = [
      { properties: { headerReferences: [{ type: "default", rId: "h-default" }] }, content: [] },
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
            content: [{ type: "run", content: [{ type: "text", text: "root" }] }],
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
            content: [{ type: "run", content: [{ type: "text", text: "needle" }] }],
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
    headerFooterRefs: number === 2 && kind === "header" ? { headerDefault: "h-default" } : {},
  })),
});

for (const change of ["secondary-story", "rendered-scope"] as const) {
  test.each(["current", "all"] as const)(
    `Vue refreshes ${change} changes before %s replacement with the same PM body`,
    (operation) => {
      const source = navigationFixture("header");
      source.package.document.content = [{ type: "paragraph", content: [textRun("hit")] }];
      const header = source.package.headers?.get("h-default");
      if (!header) throw new TypeError("Header fixture is missing");
      header.content = [{ type: "paragraph", content: [textRun("hit")] }];
      const host = document.createElement("div");
      const hidden = document.createElement("div");
      document.body.append(host, hidden);
      const view = new EditorView(hidden, {
        state: EditorState.create({ doc: toProseDoc(source) }),
      });
      const currentLayout = shallowRef(navigationLayout("header"));
      const holder: { search: UseFindReplaceReturn | null } = { search: null };
      const app = createApp(
        defineComponent({
          setup() {
            holder.search = useFindReplace({
              editorView: shallowRef(view),
              pagesRef: shallowRef(null),
              layout: currentLayout,
              getDocument: () =>
                fromProseDoc(view.state.doc, source, { stylesheetSource: { type: "package" } }),
              displayMode: ref("all-markup"),
              readOnly: ref(false),
              onResultChange: () => {},
            });
            return () => h("div");
          },
        }),
      );
      app.mount(host);
      try {
        const search = holder.search ?? panic("Search composable was not mounted");
        expect(search.find("hit")?.totalCount).toBe(2);
        const before = view.state;
        const beforeJson = before.doc.toJSON();
        if (change === "secondary-story")
          header.content = [{ type: "paragraph", content: [textRun("miss")] }];
        else {
          for (const page of currentLayout.value.pages) page.headerFooterRefs = {};
          currentLayout.value = { ...currentLayout.value };
        }
        search.scheduleSearch();
        search.replaceText.value = "changed";
        if (operation === "current") expect(search.replaceCurrent()).toBe(false);
        else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 2 });
        expect(view.state).toBe(before);
        expect(view.state.doc.toJSON()).toEqual(beforeJson);
        expect(search.getFindResult()?.totalCount).toBe(1);
        // Fresh clones from the document getter cannot keep the command stale.
        if (operation === "current") expect(search.replaceCurrent()).toBe(true);
        else expect(search.replaceAll()).toEqual({ replacedCount: 1, skippedCount: 0 });
        expect(view.state.doc.textContent).toBe("changed");
      } finally {
        app.unmount();
        view.destroy();
        host.remove();
        hidden.remove();
      }
    },
  );
}

const navigationViewport = () => {
  const container = document.createElement("div");
  container.dataset["folioScroll"] = "";
  Object.defineProperties(container, { clientHeight: { value: 100 }, clientWidth: { value: 500 } });
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
  "Vue navigation mounts an unmounted page for a %s match",
  async (kind) => {
    const fixture = navigationFixture(kind);
    const { container, shell } = navigationViewport();
    const host = document.createElement("div");
    const view = new EditorView(document.createElement("div"), {
      state: EditorState.create({ doc: toProseDoc(fixture) }),
    });
    const holder: { search: UseFindReplaceReturn | null } = { search: null };
    const scroll = spyOn(container, "scrollTo").mockImplementation(() => {
      if (!shell.querySelector("span")) queueMicrotask(() => mountNavigationMatch(shell, kind));
    });
    const app = createApp(
      defineComponent({
        setup() {
          holder.search = useFindReplace({
            editorView: shallowRef(view),
            pagesRef: shallowRef(container),
            layout: shallowRef(navigationLayout(kind)),
            getDocument: () => fixture,
            displayMode: ref("no-markup"),
            readOnly: ref(true),
            onResultChange: () => {},
          });
          return () => h("div");
        },
      }),
    );
    app.mount(host);
    try {
      expect(holder.search?.find("needle")?.totalCount).toBe(1);
      await new Promise((resolve) => setTimeout(resolve, 40));
      expect(scroll).toHaveBeenCalled();
      expect(shell.querySelector("[data-pm-start]")).not.toBeNull();
      expect(container.querySelector("[data-find-match-index='0']")).not.toBeNull();
      expect(scroll.mock.calls.length).toBeGreaterThan(1);
    } finally {
      app.unmount();
      view.destroy();
      scroll.mockRestore();
      container.remove();
    }
  },
);

test.each(["collapsed", "visible", "deferred"] as const)(
  "Vue activates a %s reply before revealing its geometry",
  async (visibility) => {
    const fixture = navigationFixture("reply");
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
    const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
      const range = createRange();
      Object.defineProperty(range, "getClientRects", {
        value: () => {
          events.push(active ? "measure-active" : "measure-inactive");
          return committed ? [new DOMRect(30, 200, 60, 18)] : [];
        },
      });
      return range;
    });
    const scroll = spyOn(container, "scrollTo").mockImplementation(() => {
      events.push(active ? "scroll-active" : "scroll-inactive");
    });
    const host = document.createElement("div");
    const view = new EditorView(document.createElement("div"), {
      state: EditorState.create({ doc: toProseDoc(fixture) }),
    });
    const holder: { search: UseFindReplaceReturn | null } = { search: null };
    const app = createApp(
      defineComponent({
        setup() {
          holder.search = useFindReplace({
            editorView: shallowRef(view),
            pagesRef: shallowRef(container),
            layout: shallowRef(null),
            getDocument: () => fixture,
            commentRoot: shallowRef(container),
            revealComment: (id) => {
              expect(id).toBe(8);
              events.push("activate");
              active = true;
              const commit = () => {
                committed = true;
                events.push("commit");
                reply.dataset["replyExpanded"] = "true";
              };
              if (visibility === "deferred") queueMicrotask(commit);
              else commit();
            },
            displayMode: ref("no-markup"),
            readOnly: ref(true),
            onResultChange: () => {},
          });
          return () => h("div");
        },
      }),
    );
    app.mount(host);
    try {
      holder.search?.find("needle");
      if (visibility === "deferred") {
        expect(events).toContain("activate");
        expect(events).not.toContain("commit");
        expect(scroll).not.toHaveBeenCalled();
      }
      await new Promise((resolve) => setTimeout(resolve, 40));
      expect(events).toContain("activate");
      expect(events).not.toContain("scroll-inactive");
      expect(
        container.querySelector("[data-folio-find-comment-overlay] [data-find-match-index='0']"),
      ).not.toBeNull();
      expect(events).toContain("scroll-active");
      expect(events.indexOf("scroll-active")).toBeGreaterThan(events.indexOf("commit"));
    } finally {
      app.unmount();
      view.destroy();
      rangeFactory.mockRestore();
      scroll.mockRestore();
      container.remove();
    }
  },
);

test("initial projected searches defer until the lazy body editor mounts", () => {
  const revisionAttrs = { revisionId: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
  const insertion = (schema.marks["insertion"] ?? panic("Missing insertion mark")).create(
    revisionAttrs,
  );
  const deletion = (schema.marks["deletion"] ?? panic("Missing deletion mark")).create(
    revisionAttrs,
  );
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("inserted", [insertion]),
        schema.text(" deleted", [deletion]),
        schema.text(" kept"),
      ]),
    ]),
  });
  const hiddenQueries = {
    "all-markup": null,
    "simple-markup": "deleted",
    "no-markup": "deleted",
    original: "inserted",
  } as const satisfies Record<DisplayMode, string | null>;
  for (const mode of DISPLAY_MODES) {
    const hiddenQuery = hiddenQueries[mode];
    if (!hiddenQuery) continue;
    const editorView = shallowRef<EditorView | null>(null);
    const scope = effectScope();
    const source = createEmptyDocument({ initialText: "inserted deleted kept" });
    const find =
      scope.run(() =>
        useFindReplace({
          editorView,
          pagesRef: shallowRef(null),
          layout: shallowRef(null),
          getDocument: () => source,
          displayMode: ref(mode),
          readOnly: ref(false),
          onResultChange: () => {},
        }),
      ) ?? panic("Expected Vue find composable");
    let view: EditorView | null = null;
    try {
      find.searchText.value = hiddenQuery;
      expect(find.performSearch()).toEqual([]);
      expect(find.findNext()).toBeNull();
      view = new EditorView(document.createElement("div"), { state });
      editorView.value = view;
      expect(find.performSearch()).toEqual([]);
      find.searchText.value = "kept";
      expect(find.performSearch()).toHaveLength(1);
    } finally {
      scope.stop();
      view?.destroy();
    }
  }
});

test("Vue searches refresh with the markup view and replacement respects read-only state", async () => {
  const deletion = (schema.marks["deletion"] ?? panic("Missing deletion mark")).create({
    revisionId: 1,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
  });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("needle", [deletion]), schema.text(" kept")]),
  ]);
  const view = new EditorView(document.createElement("div"), {
    state: EditorState.create({ doc }),
  });
  const markupView = ref<DisplayMode>("all-markup");
  const readOnly = ref(true);
  const scope = effectScope();
  const source = createEmptyDocument();
  const find =
    scope.run(() =>
      useFindReplace({
        editorView: shallowRef(view),
        pagesRef: shallowRef(null),
        layout: shallowRef(null),
        getDocument: () =>
          fromProseDoc(view.state.doc, source, { stylesheetSource: { type: "package" } }),
        displayMode: markupView,
        readOnly,
        onResultChange: () => {},
      }),
    ) ?? panic("Expected Vue find composable");
  try {
    find.searchText.value = "needle";
    expect(find.performSearch()).toHaveLength(1);
    markupView.value = "no-markup";
    await nextTick();
    expect(find.matches.value).toEqual([]);
    markupView.value = "original";
    await nextTick();
    expect(find.matches.value).toHaveLength(1);
    expect(find.currentMatch.value).toMatchObject({ authoredRange: { from: 1, to: 7 } });
    find.replaceText.value = "replacement";
    const refused = { replacedCount: 0, skippedCount: 1 };
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toEqual(refused);
    expect(view.state.doc).toBe(doc);

    readOnly.value = false;
    view.setProps({ editable: () => false });
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toEqual(refused);
    expect(view.state.doc).toBe(doc);

    view.setProps({ editable: () => true });
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toEqual(refused);
    markupView.value = "all-markup";
    await nextTick();
    expect(find.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("replacement kept");
  } finally {
    scope.stop();
    view.destroy();
  }
});

test("projected replacement refuses a match spanning hidden revisions without changing saved content", async () => {
  const input = await buildRawPackage({
    styles: WORD_STYLES,
    body:
      '<w:p><w:r><w:t>left</w:t></w:r><w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> hidden </w:delText></w:r></w:del><w:r><w:t>right</w:t></w:r></w:p>' +
      SECTION(),
  });
  const source = await parseDocx(input, { preloadFonts: false, detectVariables: false });
  const doc = toProseDoc(source);
  const view = new EditorView(document.createElement("div"), {
    state: EditorState.create({ doc }),
    editable: () => true,
  });
  const markupView = ref<DisplayMode>("no-markup");
  const reasons: EditorReadOnlyReason[] = [];
  const scope = effectScope();
  const find =
    scope.run(() =>
      useFindReplace({
        editorView: shallowRef(view),
        pagesRef: shallowRef(null),
        layout: shallowRef(null),
        getDocument: () =>
          fromProseDoc(view.state.doc, source, { stylesheetSource: { type: "package" } }),
        displayMode: markupView,
        readOnly: ref(false),
        onReadonlyEditAttempt: (reason) => reasons.push(reason),
        onResultChange: () => {},
      }),
    ) ?? panic("Expected Vue find composable");
  const originalJson = JSON.stringify(doc.toJSON());
  const serializeLiveDocument = async () =>
    new Uint8Array(
      await repackDocx(
        fromProseDoc(view.state.doc, source, { stylesheetSource: { type: "package" } }),
        { updateModifiedDate: false },
      ),
    );
  const savedBefore = await serializeLiveDocument();
  try {
    find.searchText.value = "leftright";
    expect(find.performSearch()).toMatchObject([
      { authoredRange: { from: 1, to: 18 }, text: "leftright" },
    ]);
    find.replaceText.value = "replacement";
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
    expect(reasons).toEqual(["review-display-mode", "review-display-mode"]);
    expect(JSON.stringify(view.state.doc.toJSON())).toBe(originalJson);
    const savedAfter = await serializeLiveDocument();
    expect(savedAfter).toEqual(savedBefore);
    markupView.value = "all-markup";
    await nextTick();
    find.searchText.value = "left";
    expect(find.performSearch()).toHaveLength(1);
    expect(find.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("replacement hidden right");
  } finally {
    scope.stop();
    view.destroy();
  }
});
