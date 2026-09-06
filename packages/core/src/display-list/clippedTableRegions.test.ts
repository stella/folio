import { expect, test } from "bun:test";

import { layoutDocument } from "../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../layout-engine/measure/measureBlocks";
import type { FlowBlock, PageMargins, TableBlock } from "../layout-engine/types";
import type { BlockLookup } from "../layout-painter/index";
import { buildDisplayList } from "./build/buildDisplayList";
import { renderDisplayPageToDom } from "./dom/renderDisplayListToDom";
import type { DisplayHitRegion } from "./types";

type StubElement = {
  readonly style: Record<string, string>;
  readonly dataset: Record<string, string>;
  readonly children: StubElement[];
  className: string;
  id: string;
  textContent: string;
  href: string;
  title: string;
  src: string;
  alt: string;
  readonly append: (...nodes: StubElement[]) => void;
  readonly setAttribute: (name: string, value: string) => void;
};

const createStubElement = (): StubElement => {
  const children: StubElement[] = [];
  return {
    style: {},
    dataset: {},
    children,
    className: "",
    id: "",
    textContent: "",
    href: "",
    title: "",
    src: "",
    alt: "",
    append: (...nodes) => children.push(...nodes),
    setAttribute: () => {},
  };
};

// SAFETY: the renderer touches only createElement plus the element properties
// StubElement declares; nothing here reaches the rest of the DOM API.
const stubDocument = () => ({ createElement: createStubElement }) as unknown as Document;
const asStub = (element: HTMLElement) => element as unknown as StubElement;

const findByClass = (element: StubElement, className: string): StubElement | undefined => {
  if (element.className.split(" ").includes(className)) {
    return element;
  }
  for (const child of element.children) {
    const found = findByClass(child, className);
    if (found !== undefined) {
      return found;
    }
  }
  return undefined;
};

const findAll = (
  element: StubElement,
  predicate: (candidate: StubElement) => boolean,
): StubElement[] => {
  const matches = predicate(element) ? [element] : [];
  return matches.concat(element.children.flatMap((child) => findAll(child, predicate)));
};

const flattenRegions = (regions: readonly DisplayHitRegion[]): DisplayHitRegion[] =>
  regions.flatMap((region) => [region, ...flattenRegions(region.children)]);

test("clipped table-row continuations keep editing regions", () => {
  withFakeTextMeasure(
    () => {
      const margins: PageMargins = { top: 10, right: 10, bottom: 10, left: 10 };
      const table: TableBlock = {
        kind: "table",
        id: "table",
        columnWidths: [80],
        rows: [
          {
            id: "header-row",
            isHeader: true,
            cells: [
              {
                id: "header-cell",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [
                  {
                    kind: "paragraph",
                    id: "header-paragraph",
                    runs: [{ kind: "text", text: "Header", pmStart: 1, pmEnd: 7 }],
                    pmStart: 0,
                    pmEnd: 8,
                  },
                ],
              },
            ],
          },
          {
            id: "body-row",
            cells: [
              {
                id: "body-cell",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [
                  {
                    kind: "paragraph",
                    id: "body-paragraph",
                    runs: [
                      {
                        kind: "text",
                        text: "body ".repeat(30),
                        pmStart: 10,
                        pmEnd: 160,
                      },
                    ],
                    pmStart: 9,
                    pmEnd: 161,
                  },
                ],
              },
            ],
          },
        ],
      };
      const blocks: FlowBlock[] = [table];
      const measures = measureBlocks(blocks, 80);
      const blockLookup: BlockLookup = new Map([
        [String(table.id), { block: table, measure: measures[0]! }],
      ]);
      const list = buildDisplayList({
        layout: layoutDocument(blocks, measures, { pageSize: { w: 100, h: 100 }, margins }),
        blockLookup,
      });
      const pageIndex = list.pages.findIndex(
        (page, index) =>
          index > 0 && page.primitives.some((primitive) => primitive.kind === "clipGroup"),
      );
      expect(pageIndex).toBeGreaterThan(0);
      const page = list.pages[pageIndex];
      expect(page).toBeDefined();
      if (page === undefined) {
        return;
      }
      const clipPrimitive = page.primitives.find((primitive) => primitive.kind === "clipGroup");
      expect(clipPrimitive?.kind).toBe("clipGroup");
      if (clipPrimitive?.kind !== "clipGroup") {
        return;
      }

      const pageRegions = flattenRegions(page.regions);
      const clipRegions = flattenRegions(clipPrimitive.regions);
      const tableRegion = pageRegions.find((region) => region.kind === "table");
      const rowRegion = clipPrimitive.regions.find(
        (region) => region.kind === "tableRow" && region.model?.rowIndex === 1,
      );
      const paragraphRegion = clipRegions.find(
        (region) => region.kind === "paragraph" && region.model?.blockId === "body-paragraph",
      );
      const lineRegion = clipRegions.find(
        (region) => region.kind === "line" && region.model?.blockId === "body-paragraph",
      );
      expect(tableRegion).toBeDefined();
      expect(rowRegion).toBeDefined();
      expect(paragraphRegion).toBeDefined();
      expect(lineRegion?.model?.pmRange).toBeDefined();
      if (
        tableRegion === undefined ||
        rowRegion === undefined ||
        paragraphRegion === undefined ||
        lineRegion?.model?.pmRange === undefined
      ) {
        return;
      }

      const rendered = asStub(
        renderDisplayPageToDom(page, {
          doc: stubDocument(),
          fonts: list.fonts,
          images: list.images,
          pageIndex,
        }),
      );
      const pageContent = findByClass(rendered, "layout-page-content");
      const tableElement = pageContent && findByClass(pageContent, "layout-table");
      expect(tableElement).toBeDefined();
      if (tableElement === undefined) {
        return;
      }
      const clips = findAll(
        tableElement,
        (element) => element !== tableElement && element.style.overflow === "hidden",
      );
      expect(clips).toHaveLength(1);
      const clip = clips[0];
      expect(clip).toBeDefined();
      if (clip === undefined) {
        return;
      }
      expect(clip.style.left).toBe(`${clipPrimitive.rect.xPx - tableRegion.rect.xPx}px`);
      expect(clip.style.top).toBe(`${clipPrimitive.rect.yPx - tableRegion.rect.yPx}px`);

      const row = findByClass(clip, "layout-table-row");
      expect(row).toBeDefined();
      if (row === undefined) {
        return;
      }
      expect(row.dataset).toMatchObject({ rowIndex: "1" });
      expect(row.style.left).toBe(`${rowRegion.rect.xPx - clipPrimitive.rect.xPx}px`);
      expect(row.style.top).toBe(`${rowRegion.rect.yPx - clipPrimitive.rect.yPx}px`);
      expect(findByClass(clip, "layout-table-cell")?.dataset).toMatchObject({
        rowIndex: "1",
        columnIndex: "0",
      });
      expect(findByClass(clip, "layout-paragraph")?.dataset).toMatchObject({
        blockId: "body-paragraph",
      });
      const line = findByClass(clip, "layout-line");
      expect(line).toBeDefined();
      if (line === undefined) {
        return;
      }
      expect(line.dataset).toMatchObject({
        blockId: "body-paragraph",
        pmStart: String(lineRegion.model.pmRange.start),
        pmEnd: String(lineRegion.model.pmRange.end),
        story: "body",
      });
      expect(line.style.left).toBe(`${lineRegion.rect.xPx - paragraphRegion.rect.xPx}px`);
      expect(line.style.top).toBe(`${lineRegion.rect.yPx - paragraphRegion.rect.yPx}px`);

      const headerRows = findAll(
        rendered,
        (element) =>
          element.className.split(" ").includes("layout-table-row") &&
          element.dataset["rowIndex"] === "0",
      );
      expect(headerRows).toHaveLength(1);
      expect(findByClass(clip, "layout-paragraph")?.dataset["blockId"]).not.toBe(
        "header-paragraph",
      );
    },
    { charWidth: fixedCharWidth(5) },
  );
});
