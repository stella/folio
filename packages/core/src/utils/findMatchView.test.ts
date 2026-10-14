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
          : headerFooterToProseDoc(storyBody(doc, range.story).content, {
              ...(doc.package.styles === undefined ? {} : { styles: doc.package.styles }),
              ...(doc.package.theme === undefined ? {} : { theme: doc.package.theme }),
              ...(doc.package.numbering === undefined ? {} : { numbering: doc.package.numbering }),
            });
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
  test.each(["footnoteRef", "endnoteRef"] as const)(
    "%s labels inside a combined painted span are excluded from search decoration",
    (type) => {
      const doc: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "A" },
                      { type, id: 12 },
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const before = structuredClone(doc);
      const bodyDoc = toProseDoc(doc);
      expect(bodyDoc.textContent).toBe("A12B");
      const matches = findInDocument(doc, "AB", createDefaultFindOptions());
      expect(matches).toHaveLength(1);
      expect(matches.at(0)).toMatchObject({
        startOffset: 0,
        endOffset: 2,
        text: "AB",
        replacement: { status: "blocked", reason: "atomic-content" },
      });
      expect(resolveDocumentFindRanges({ document: doc, bodyDoc, matches })).toEqual([
        { from: 1, to: 5, index: 0, story: "main", excludedRanges: [{ from: 2, to: 4 }] },
      ]);
      const container = viewport();
      const pageElement = document.createElement("div");
      pageElement.className = "layout-page";
      const span = document.createElement("span");
      span.dataset["pmStart"] = "1";
      span.dataset["pmEnd"] = "5";
      span.textContent = "A12B";
      span.getBoundingClientRect = () => new DOMRect(30, 200, 40, 18);
      pageElement.append(span);
      container.append(pageElement);
      const intervals: { start: number; end: number; text: string }[] = [];
      const createRange = document.createRange.bind(document);
      const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
        const range = createRange();
        Object.defineProperty(range, "getClientRects", {
          value: () => {
            intervals.push({
              start: range.startOffset,
              end: range.endOffset,
              text: range.toString(),
            });
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
      const scroll = spyOn(container, "scrollTo").mockImplementation(() => {});
      const view = createFindMatchView({ container, document: doc, bodyDoc, matches });
      expect(intervals).toEqual([
        { start: 0, end: 1, text: "A" },
        { start: 3, end: 4, text: "B" },
      ]);
      const marks = [...container.querySelectorAll<HTMLElement>("[data-find-match-index]")];
      expect(marks.map((mark) => ({ left: mark.style.left, width: mark.style.width }))).toEqual([
        { left: "30px", width: "10px" },
        { left: "60px", width: "10px" },
      ]);
      expect(marks.every((mark) => mark.dataset["current"] === "true")).toBe(true);
      view.paint(-1);
      expect(marks.every((mark) => mark.dataset["current"] === "false")).toBe(true);
      view.paint(0);
      expect(marks.every((mark) => mark.dataset["current"] === "true")).toBe(true);
      expect(view.reveal(0)).toBe(true);
      expect(scroll).toHaveBeenCalledWith({ top: 159, left: 0, behavior: "smooth" });
      expect(doc).toEqual(before);
      expect(span.textContent).toBe("A12B");
      view.destroy();
      rangeFactory.mockRestore();
      scroll.mockRestore();
    },
  );

  test("cached field substrings and following text use separate painted coordinates", () => {
    const doc: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "complexField",
                  fieldType: "DOCVARIABLE",
                  instruction: " DOCVARIABLE label ",
                  fieldCode: [],
                  fieldResult: [
                    { type: "run", content: [{ type: "text", text: "field hit value" }] },
                  ],
                },
                { type: "run", content: [{ type: "text", text: " after hit" }] },
              ],
            },
          ],
        },
      },
    };
    const before = structuredClone(doc);
    const bodyDoc = toProseDoc(doc);
    expect(bodyDoc.child(0).child(0).type.name).toBe("field");
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    expect(matches).toHaveLength(2);
    const ranges = resolveDocumentFindRanges({ document: doc, bodyDoc, matches });
    expect(ranges).toEqual([
      {
        from: 1,
        to: 2,
        index: 0,
        story: "main",
        atomicTextRanges: [{ from: 1, to: 2, startOffset: 6, endOffset: 9 }],
      },
      { from: 9, to: 12, index: 1, story: "main" },
    ]);
    const container = viewport();
    const pageElement = document.createElement("div");
    pageElement.className = "layout-page";
    const fieldSpan = document.createElement("span");
    fieldSpan.dataset["decorationStart"] = "1";
    fieldSpan.dataset["decorationEnd"] = "2";
    fieldSpan.textContent = "field hit value";
    const followingSpan = document.createElement("span");
    followingSpan.dataset["pmStart"] = "2";
    followingSpan.dataset["pmEnd"] = "12";
    followingSpan.textContent = " after hit";
    pageElement.append(fieldSpan, followingSpan);
    container.append(pageElement);
    const measured: { text: string | null; start: number; end: number }[] = [];
    const createRange = document.createRange.bind(document);
    const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
      const range = createRange();
      Object.defineProperty(range, "getClientRects", {
        value: () => {
          measured.push({
            text: range.startContainer.textContent,
            start: range.startOffset,
            end: range.endOffset,
          });
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
    const view = createFindMatchView({ container, document: doc, bodyDoc, matches });
    expect(measured).toEqual([
      { text: "field hit value", start: 6, end: 9 },
      { text: " after hit", start: 7, end: 10 },
    ]);
    expect(
      [...container.querySelectorAll<HTMLElement>("[data-find-match-index]")].map(
        (mark) => mark.style.width,
      ),
    ).toEqual(["30px", "30px"]);
    expect(fieldSpan.hasAttribute("data-pm-start")).toBe(false);
    expect(fieldSpan.hasAttribute("data-pm-end")).toBe(false);
    expect(doc).toEqual(before);
    view.destroy();
    rangeFactory.mockRestore();
  });

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
    const commentScroll = spyOn(commentRoot, "scrollTo").mockImplementation(() => {});
    const bodyScroll = spyOn(container, "scrollTo").mockImplementation(() => {});
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
    expect(commentScroll).toHaveBeenCalledWith({ top: 219, left: 0, behavior: "smooth" });
    expect(bodyScroll).not.toHaveBeenCalled();
    expect(doc).toEqual(before);
    expect(commentText.textContent).toBe("comment hit");
    view.destroy();
    expect(document.querySelectorAll("[data-find-match-index]").length).toBe(0);
    rangeFactory.mockRestore();
    commentScroll.mockRestore();
    bodyScroll.mockRestore();
  });

  test.each(["read-only", "editing"])(
    "%s highlights every story and moves current emphasis without editing",
    (mode) => {
      const doc = fixture();
      const before = structuredClone(doc);
      const matches = findInDocument(doc, "hit", createDefaultFindOptions());
      const container = viewport();
      const scroll = spyOn(container, "scrollTo").mockImplementation(() => {});
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
      const currentMark = overlay?.querySelector<HTMLElement>('[data-current="true"]');
      expect(currentMark?.style.outlineWidth).toBe("2px");
      expect(currentMark?.style.outlineStyle).toBe("solid");
      expect(
        overlay?.querySelector<HTMLElement>('[data-current="false"]')?.style.outlineWidth,
      ).toBe("0px");
      for (const [index] of matches.entries()) {
        scroll.mockClear();
        expect(view.reveal(index)).toBe(true);
        expect(scroll).toHaveBeenCalledWith({ top: 159 + index * 40, left: 0, behavior: "smooth" });
      }
      expect(view.reveal(matches.length)).toBe(false);
      expect(
        [...container.querySelectorAll(".layout-page")].map((element) => element.outerHTML),
      ).toEqual(paintedBefore);
      expect(doc).toEqual(before);
      view.destroy();
      expect(container.querySelector("[data-folio-find-overlay]")).toBeNull();
      scroll.mockRestore();
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
    const scroll = spyOn(container, "scrollTo").mockImplementation(() => {});
    appendPaintedSpan({ container, ...range, from: 1, to: 18, top: 200 });
    appendPaintedSpan({ container, ...range, from: 1, to: 7, top: 800 });
    appendPaintedSpan({ container, ...range, from: 7, to: range.from + 1, top: 800 });
    appendPaintedSpan({ container, ...range, from: range.from + 1, top: 800 });
    appendPaintedSpan({ container, ...range, from: range.to, to: 15, top: 800 });
    appendPaintedSpan({ container, ...range, from: 15, to: 18, top: 800 });
    const view = createFindMatchView({ container, document: doc, matches });
    expect(container.querySelectorAll("[data-find-match-index]").length).toBe(3);
    expect(view.reveal(0)).toBe(true);
    expect(scroll).toHaveBeenCalledWith({ top: 159, left: 0, behavior: "smooth" });
    view.destroy();
    scroll.mockRestore();
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

// Navigation needs browser visibility, not merely a text node in the DOM.
// Exercise the geometry oracle with both empty and zero-area ranges.
test.each(
  (["comment", "body"] as const).flatMap((surface) =>
    (["empty", "zero-width", "zero-height"] as const).map((visibility) => ({
      surface,
      visibility,
    })),
  ),
)(
  "hidden $surface range ($visibility) cannot be revealed until it becomes visible",
  ({ surface, visibility }) => {
    const doc = {
      package: {
        document: {
          content: [paragraph("body")],
          comments: [{ id: 8, parentId: 7, author: "Author", content: [paragraph("reply hit")] }],
        },
      },
    } satisfies Document;
    const container = viewport();
    const reply = document.createElement("span");
    reply.dataset["findCommentId"] = "8";
    reply.dataset["findCommentParagraph"] = "0";
    reply.textContent = "reply hit";
    container.append(reply);
    let visible = false;
    const createRange = document.createRange.bind(document);
    const rangeFactory = spyOn(document, "createRange").mockImplementation(() => {
      const range = createRange();
      Object.defineProperty(range, "getClientRects", {
        value: () => {
          if (visible) return [new DOMRect(20, 200, 30, 18)];
          if (visibility === "empty") return [];
          return [
            new DOMRect(
              20,
              200,
              visibility === "zero-width" ? 0 : 30,
              visibility === "zero-height" ? 0 : 18,
            ),
          ];
        },
      });
      return range;
    });
    if (surface === "body") {
      doc.package.document.content = [paragraph("reply hit")];
      const pageElement = document.createElement("div");
      pageElement.className = "layout-page";
      delete reply.dataset["findCommentId"];
      delete reply.dataset["findCommentParagraph"];
      reply.dataset["pmStart"] = "1";
      reply.dataset["pmEnd"] = "10";
      pageElement.append(reply);
      container.append(pageElement);
    }
    const matches = findInDocument(doc, "hit", {
      ...createDefaultFindOptions(),
      includeComments: surface === "comment",
    });
    const view = createFindMatchView({ container, document: doc, matches });
    try {
      expect(view.reveal(0)).toBe(false);
      expect(container.querySelectorAll("[data-find-match-index]")).toHaveLength(0);
      visible = true;
      expect(view.reveal(0)).toBe(true);
      expect(container.querySelectorAll("[data-find-match-index]")).toHaveLength(1);
    } finally {
      view.destroy();
      rangeFactory.mockRestore();
    }
  },
);
