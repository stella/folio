import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, afterEach, describe, expect, spyOn, test } from "bun:test";
import { OP_STORIES, sameStory, storyBody } from "@stll/docx-core/ops";
import { Schema } from "prosemirror-model";

import type { Layout, Page } from "../layout-engine/types";
import { headerFooterToProseDoc, toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Paragraph } from "../types/content";
import type { Document } from "../types/document";
import {
  createFindMatchView,
  getRenderedSearchStories,
  resolveDocumentFindRanges,
} from "./findMatchView";
import { createDefaultFindOptions, findInDocument } from "./findReplace";

afterAll(() => GlobalRegistrator.unregister());
afterEach(() => {
  document.body.replaceChildren();
});

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const fixture = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("body hit"),
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    size: { width: 914400, height: 914400 },
                    textBody: { content: [paragraph("box hit")] },
                  },
                },
              ],
            },
          ],
        },
      ],
      sections: [
        {
          properties: {
            headerReferences: (["default", "first", "even"] as const).map((type) => ({
              type,
              rId: `h-${type}`,
            })),
            footerReferences: (["default", "first", "even"] as const).map((type) => ({
              type,
              rId: `f-${type}`,
            })),
          },
          content: [],
        },
      ],
    },
    headers: new Map([
      [
        "h-default",
        { type: "header", hdrFtrType: "default", content: [paragraph("header hit suffix")] },
      ],
      ["h-first", { type: "header", hdrFtrType: "first", content: [paragraph("header hit")] }],
      ["h-even", { type: "header", hdrFtrType: "even", content: [paragraph("header hit")] }],
    ]),
    footers: new Map([
      ["f-default", { type: "footer", hdrFtrType: "default", content: [paragraph("footer hit")] }],
      ["f-first", { type: "footer", hdrFtrType: "first", content: [paragraph("footer hit")] }],
      ["f-even", { type: "footer", hdrFtrType: "even", content: [paragraph("footer hit")] }],
    ]),
    footnotes: [{ type: "footnote", id: 1, content: [paragraph("footnote hit")] }],
    endnotes: [{ type: "endnote", id: 2, content: [paragraph("endnote hit")] }],
  },
});

const page = (number: number, refs: NonNullable<Page["headerFooterRefs"]>): Page => ({
  number,
  logicalNumber: number,
  sectionPageNumber: number,
  fragments: [],
  margins: { top: 20, right: 20, bottom: 20, left: 20 },
  size: { w: 800, h: 1000 },
  headerFooterRefs: refs,
});

const layout = (pages: Page[]): Layout => ({ pageSize: { w: 800, h: 1000 }, pages });

describe("rendered search story eligibility", () => {
  const refs = {
    headerDefault: "h-default",
    headerFirst: "h-first",
    headerEven: "h-even",
    footerDefault: "f-default",
    footerFirst: "f-first",
    footerEven: "f-even",
  };

  test("inactive first/even variants are excluded and shared parts count once", () => {
    const stories = getRenderedSearchStories(layout([page(1, refs), page(2, refs), page(3, refs)]));
    expect(stories).toEqual([
      { kind: "header", rId: "h-default" },
      { kind: "footer", rId: "f-default" },
    ]);
    const matches = findInDocument(fixture(), "hit", createDefaultFindOptions(), {
      renderedStories: stories,
    });
    expect(matches).toHaveLength(6);
    expect(
      matches.filter(({ story }) => story !== OP_STORIES.MAIN).map(({ story }) => story),
    ).toEqual([
      { kind: "header", rId: "h-default" },
      { kind: "footer", rId: "f-default" },
      { kind: "footnote", id: 1 },
      { kind: "endnote", id: 2 },
    ]);
  });

  test("first-page precedence and enabled even variants follow rendered pages", () => {
    const enabled = { ...refs, titlePg: true, evenAndOddHeaders: true };
    const first = { ...page(2, enabled), sectionPageNumber: 1 };
    expect(getRenderedSearchStories(layout([first]))).toEqual([
      { kind: "header", rId: "h-first" },
      { kind: "footer", rId: "f-first" },
    ]);
    const stories = getRenderedSearchStories(
      layout([page(1, enabled), page(2, enabled), page(3, enabled)]),
    );
    expect(stories).toHaveLength(6);
    expect(
      findInDocument(fixture(), "hit", createDefaultFindOptions(), { renderedStories: stories }),
    ).toHaveLength(10);
    expect(getRenderedSearchStories(null)).toEqual([]);
  });

  test("an active variant with no part does not fall back to hidden content", () => {
    expect(
      getRenderedSearchStories(
        layout([
          page(1, { titlePg: true, headerDefault: "h-default", footerDefault: "f-default" }),
        ]),
      ),
    ).toEqual([]);
  });
});

describe("story-local painted search ranges", () => {
  test("all body, shape, variant and note matches resolve to their own story text", () => {
    const doc = fixture();
    const before = structuredClone(doc);
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    const ranges = resolveDocumentFindRanges({ document: doc, matches });
    expect(ranges).toHaveLength(matches.length);
    expect(new Set(ranges.map(({ index }) => index)).size).toBe(matches.length);
    for (const range of ranges) {
      const match = matches.at(range.index);
      expect(match).toBeDefined();
      expect(sameStory(match?.story ?? OP_STORIES.MAIN, range.story)).toBe(true);
      const source =
        range.story === OP_STORIES.MAIN
          ? toProseDoc(doc)
          : headerFooterToProseDoc(storyBody(doc, range.story).content, doc.package);
      expect(source.textBetween(range.from, range.to)).toBe("hit");
    }
    expect(doc).toEqual(before);
  });

  test("live body positions include shape offsets without confusing body paragraphs", () => {
    const schema = new Schema({
      nodes: {
        doc: { content: "(paragraph | textBox)+" },
        paragraph: { content: "text*" },
        textBox: { content: "paragraph+" },
        text: {},
      },
    });
    const bodyDoc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("body hit")]),
      schema.node("textBox", null, [schema.node("paragraph", null, [schema.text("box hit")])]),
    ]);
    const doc = fixture();
    const matches = findInDocument(doc, "hit", createDefaultFindOptions()).filter(
      ({ story }) => story === OP_STORIES.MAIN,
    );
    const ranges = resolveDocumentFindRanges({ document: doc, bodyDoc, matches });
    expect(ranges.map(({ from, to }) => bodyDoc.textBetween(from, to))).toEqual(["hit", "hit"]);
    expect(ranges.map(({ from, to }) => ({ from, to }))).toEqual([
      { from: 6, to: 9 },
      { from: 16, to: 19 },
    ]);
  });
});

// Geometry comes from the browser in production. Tabs expose deterministic element
// geometry here, keeping the overlay assertions independent of a DOM text shaper.
const appendPaintedSpan = ({
  container,
  from,
  to,
  top,
  story,
}: {
  container: HTMLElement;
  from: number;
  to: number;
  top: number;
  story: ReturnType<typeof resolveDocumentFindRanges>[number]["story"];
}) => {
  const pageElement = document.createElement("div");
  pageElement.className = "layout-page";
  const span = document.createElement("span");
  span.className = "layout-run-tab";
  span.dataset["pmStart"] = String(from);
  span.dataset["pmEnd"] = String(to);
  if (story !== OP_STORIES.MAIN) {
    switch (story.kind) {
      case "header":
        span.dataset["story"] = story.kind;
        span.dataset["hfRid"] = story.rId;
        break;
      case "footer":
        pageElement.dataset["hfSlotKind"] = story.kind;
        pageElement.dataset["hfRid"] = story.rId;
        break;
      case "footnote":
        span.dataset["story"] = story.kind;
        span.dataset["noteId"] = String(story.id);
        break;
      case "endnote":
        pageElement.dataset["noteKind"] = story.kind;
        pageElement.dataset["noteId"] = String(story.id);
        break;
      default: {
        const exhaustive: never = story;
        return exhaustive;
      }
    }
  }
  span.getBoundingClientRect = () => new DOMRect(30, top, 24, 18);
  pageElement.append(span);
  container.append(pageElement);
  return span;
};

const viewport = () => {
  const container = document.createElement("div");
  container.dataset["folioScroll"] = "";
  Object.defineProperties(container, {
    clientHeight: { value: 100 },
    clientWidth: { value: 500 },
  });
  container.getBoundingClientRect = () => new DOMRect(0, 0, 500, 100);
  document.body.append(container);
  return container;
};

describe("view-only search geometry", () => {
  test.each([0.5, 2])(
    "scaled paint surface keeps search marks in local coordinates at zoom %s",
    (zoom) => {
      const doc: Document = { package: { document: { content: [paragraph("hit")] } } };
      const container = viewport();
      const span = appendPaintedSpan({
        container,
        from: 1,
        to: 4,
        top: 100,
        story: OP_STORIES.MAIN,
      });
      span.getBoundingClientRect = () =>
        new DOMRect(10 + 40 * zoom, 20 + 80 * zoom, 24 * zoom, 18 * zoom);
      const view = createFindMatchView({
        container,
        document: doc,
        matches: findInDocument(doc, "hit", createDefaultFindOptions()),
      });
      const overlay = container.querySelector<HTMLElement>("[data-folio-find-overlay]");
      expect(overlay).not.toBeNull();
      if (!overlay) return;
      Object.defineProperties(overlay, {
        offsetWidth: { value: 500 },
        offsetHeight: { value: 100 },
      });
      overlay.getBoundingClientRect = () => new DOMRect(10, 20, 500 * zoom, 100 * zoom);
      // A missing target forces a fresh paint after installing browser geometry.
      expect(view.reveal(1)).toBe(false);
      const mark = overlay.querySelector<HTMLElement>("[data-find-match-index]");
      expect({
        left: mark?.style.left,
        top: mark?.style.top,
        width: mark?.style.width,
        height: mark?.style.height,
      }).toEqual({
        left: "40px",
        top: "80px",
        width: "24px",
        height: "18px",
      });
      view.destroy();
    },
  );

  test("each text match measures only its character interval", () => {
    const doc: Document = { package: { document: { content: [paragraph("hit hit")] } } };
    const container = viewport();
    const pageElement = document.createElement("div");
    pageElement.className = "layout-page";
    const span = document.createElement("span");
    span.dataset["pmStart"] = "1";
    span.dataset["pmEnd"] = "8";
    span.textContent = "hit hit";
    pageElement.append(span);
    container.append(pageElement);
    const createRange = document.createRange.bind(document);
    const intervals: { start: number; end: number }[] = [];
    const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
      const range = createRange();
      Object.defineProperty(range, "getClientRects", {
        value: () => {
          intervals.push({ start: range.startOffset, end: range.endOffset });
          return [
            new DOMRect(
              30 + range.startOffset * 10,
              200,
              (range.endOffset - range.startOffset) * 10,
              18,
            ),
          ];
        },
      });
      return range;
    });
    const view = createFindMatchView({
      container,
      document: doc,
      matches: findInDocument(doc, "hit", createDefaultFindOptions()),
    });
    expect(intervals).toEqual([
      { start: 0, end: 3 },
      { start: 4, end: 7 },
    ]);
    const marks = [...container.querySelectorAll<HTMLElement>("[data-find-match-index]")];
    expect(marks.map((mark) => ({ left: mark.style.left, width: mark.style.width }))).toEqual([
      { left: "30px", width: "30px" },
      { left: "70px", width: "30px" },
    ]);
    expect(span.textContent).toBe("hit hit");
    view.destroy();
    rangeFactory.mockRestore();
  });

  test("comment text outside the page surface shares the search cursor and geometry owner", () => {
    const doc: Document = {
      package: {
        document: {
          content: [paragraph("body hit")],
          comments: [{ id: 7, author: "Author", content: [paragraph("comment hit")] }],
        },
      },
    };
    const before = structuredClone(doc);
    const container = viewport();
    appendPaintedSpan({ container, from: 6, to: 9, top: 200, story: OP_STORIES.MAIN });
    const commentRoot = document.createElement("div");
    commentRoot.dataset["folioScroll"] = "";
    Object.defineProperties(commentRoot, {
      clientHeight: { value: 100 },
      clientWidth: { value: 500 },
    });
    commentRoot.getBoundingClientRect = () => new DOMRect(0, 0, 500, 100);
    const commentText = document.createElement("span");
    commentText.dataset["findCommentId"] = "7";
    commentText.dataset["findCommentParagraph"] = "0";
    commentText.textContent = "comment hit";
    commentText.getBoundingClientRect = () => new DOMRect(80, 260, 110, 18);
    commentRoot.append(commentText);
    document.body.append(commentRoot);
    const intervals: { start: number; end: number }[] = [];
    const createRange = document.createRange.bind(document);
    const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
      const range = createRange();
      Object.defineProperty(range, "getClientRects", {
        value: () => {
          intervals.push({ start: range.startOffset, end: range.endOffset });
          return [new DOMRect(160, 260, 30, 18)];
        },
      });
      return range;
    });
    const matches = findInDocument(doc, "hit", {
      ...createDefaultFindOptions(),
      includeComments: true,
    });
    const view = createFindMatchView({ container, document: doc, matches, commentRoot });
    expect(intervals).toEqual([{ start: 8, end: 11 }]);
    expect(document.querySelectorAll("[data-find-match-index]").length).toBe(2);
    view.paint(1);
    expect(
      document.querySelector('[data-current="true"]')?.getAttribute("data-find-match-index"),
    ).toBe("1");
    expect(view.reveal(1)).toBe(true);
    expect(commentRoot.scrollTop).toBe(219);
    expect(container.scrollTop).toBe(0);
    expect(doc).toEqual(before);
    expect(commentText.textContent).toBe("comment hit");
    view.destroy();
    expect(document.querySelectorAll("[data-find-match-index]").length).toBe(0);
    rangeFactory.mockRestore();
  });

  test.each(["read-only", "editing"])(
    "%s highlights every story and moves current emphasis without editing",
    (mode) => {
      const doc = fixture();
      const before = structuredClone(doc);
      const matches = findInDocument(doc, "hit", createDefaultFindOptions());
      const container = viewport();
      container.contentEditable = String(mode === "editing");
      for (const range of resolveDocumentFindRanges({ document: doc, matches })) {
        appendPaintedSpan({ container, ...range, top: 200 + range.index * 40 });
      }
      const paintedBefore = [...container.querySelectorAll(".layout-page")].map(
        (element) => element.outerHTML,
      );
      const view = createFindMatchView({ container, document: doc, matches });
      const overlay = container.querySelector<HTMLElement>("[data-folio-find-overlay]");
      expect(overlay?.getAttribute("aria-hidden")).toBe("true");
      expect(overlay?.style.pointerEvents).toBe("none");
      expect(overlay?.children.length).toBe(matches.length);
      expect(overlay?.querySelectorAll('[data-current="true"]').length).toBe(1);
      view.paint(matches.length - 1);
      expect(
        overlay?.querySelector('[data-current="true"]')?.getAttribute("data-find-match-index"),
      ).toBe(String(matches.length - 1));
      expect(overlay?.querySelector<HTMLElement>('[data-current="true"]')?.style.outline).toContain(
        "2px solid",
      );
      expect(overlay?.querySelector<HTMLElement>('[data-current="false"]')?.style.outline).toBe(
        "none",
      );
      for (const [index] of matches.entries()) {
        container.scrollTop = 0;
        expect(view.reveal(index)).toBe(true);
        expect(container.scrollTop).toBe(159 + index * 40);
      }
      expect(view.reveal(matches.length)).toBe(false);
      expect(
        [...container.querySelectorAll(".layout-page")].map((element) => element.outerHTML),
      ).toEqual(paintedBefore);
      expect(doc).toEqual(before);
      view.destroy();
      expect(container.querySelector("[data-folio-find-overlay]")).toBeNull();
    },
  );

  test("shared-part navigation selects the first rendered occurrence despite different run segmentation", () => {
    const doc = fixture();
    const matches = findInDocument(doc, "hit", createDefaultFindOptions()).filter(
      ({ story }) =>
        story !== OP_STORIES.MAIN && story?.kind === "header" && story.rId === "h-default",
    );
    const range = resolveDocumentFindRanges({ document: doc, matches }).at(0);
    expect(range).toBeDefined();
    if (!range) return;
    const container = viewport();
    appendPaintedSpan({ container, ...range, from: 1, to: 18, top: 200 });
    appendPaintedSpan({ container, ...range, from: 1, to: 7, top: 800 });
    appendPaintedSpan({ container, ...range, from: 7, to: range.from + 1, top: 800 });
    appendPaintedSpan({ container, ...range, from: range.from + 1, top: 800 });
    appendPaintedSpan({ container, ...range, from: range.to, to: 15, top: 800 });
    appendPaintedSpan({ container, ...range, from: 15, to: 18, top: 800 });
    const view = createFindMatchView({ container, document: doc, matches });
    expect(container.querySelectorAll("[data-find-match-index]").length).toBe(3);
    expect(view.reveal(0)).toBe(true);
    expect(container.scrollTop).toBe(159);
    view.destroy();
  });

  test("destroy disconnects both observers", () => {
    const mutationDisconnect = spyOn(MutationObserver.prototype, "disconnect");
    const resizeDisconnect = spyOn(ResizeObserver.prototype, "disconnect");
    const doc: Document = { package: { document: { content: [paragraph("hit")] } } };
    const container = viewport();
    const view = createFindMatchView({
      container,
      document: doc,
      matches: findInDocument(doc, "hit", createDefaultFindOptions()),
    });
    view.destroy();
    expect(mutationDisconnect).toHaveBeenCalledTimes(1);
    expect(resizeDisconnect).toHaveBeenCalledTimes(1);
    mutationDisconnect.mockRestore();
    resizeDisconnect.mockRestore();
  });
});
