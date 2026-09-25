import { describe, expect, test } from "bun:test";

import { buildDocxDisplayList, exportDocxToPdf } from "../../export-pdf";
import type { HeadlessFontSource } from "../../fonts/headlessMeasure";
import { selectDisplayPages } from "../selectDisplayPages";
import { renderDisplayListPagesToHtml, renderDisplayListToHtml } from "./renderDisplayListToHtml";

const DEMO = new URL("../../../../../tests/visual/fixtures/docx-editor-demo.docx", import.meta.url);
const NO_FONTS: HeadlessFontSource = { load: () => [] };

const demoList = async () => {
  const built = await buildDocxDisplayList(await Bun.file(DEMO).arrayBuffer(), { fonts: NO_FONTS });
  if (built.isErr()) throw built.error;
  return built.value.list;
};

describe("renderDisplayListToHtml", () => {
  test("serializes one page slot per page with the DOM backend's page markup", async () => {
    const list = await demoList();

    const pages = renderDisplayListPagesToHtml(list);
    const html = renderDisplayListToHtml(list, { title: "a <title>", pageGapPx: 16 });

    expect(pages.length).toBe(list.pages.length);
    expect(pages.every((page) => page.includes('class="layout-page"'))).toBe(true);
    expect(html.match(/class="layout-page-slot"/gu)?.length).toBe(list.pages.length);
    expect(html).toContain("<title>a &lt;title&gt;</title>");
    expect(html).not.toContain("blob:");
  });

  test("keeps authored CSS text from closing the style element", async () => {
    const list = await demoList();

    const html = renderDisplayListToHtml(list, {
      fontFaceCss: "/* </style><script>x</script> */",
      canvasColor: "red}</style><script>y</script>",
    });

    expect(html).not.toContain("</style><script>");
  });
});

describe("selectDisplayPages", () => {
  test("keeps the chosen pages and re-indexes what points at pages", async () => {
    const list = await demoList();
    const last = list.pages.length - 1;

    const selected = selectDisplayPages(list, [last, last, 0, 999]);

    expect(selected.pages.length).toBe(Math.min(2, list.pages.length));
    expect(selected.pages[0]).toEqual(
      expect.objectContaining({ widthPx: list.pages[last]?.widthPx }),
    );
    for (const page of selected.pages) {
      for (const link of page.links) {
        if (link.target.kind === "page") {
          expect(link.target.pageIndex).toBeLessThan(selected.pages.length);
        }
      }
    }
    for (const entry of selected.outline) {
      expect(entry.pageIndex).toBeLessThan(selected.pages.length);
    }
  });

  test("exportDocxToPdf writes only the requested pages", async () => {
    const bytes = await Bun.file(DEMO).arrayBuffer();

    const one = await exportDocxToPdf(bytes, {
      fonts: NO_FONTS,
      timestamp: "2026-01-01T00:00:00.000Z",
      pages: [0],
    });

    expect(one.isOk() && one.value.pageCount).toBe(1);
  });
});
