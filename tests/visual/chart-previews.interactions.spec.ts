import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** 480 by 288 CSS pixels. */
const CHART_EXTENT = 'cx="4572000" cy="2743200"';
const CHART_HEIGHT_PX = 288;
/** The series colour: nothing else on the page uses it. */
const BAR_HEX = "C0392B";
const BAR_RGB = [0xc0, 0x39, 0x2b] as const;

const paragraph = (text: string): string =>
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const chartParagraph = (rId: string, id: number): string =>
  `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent ${CHART_EXTENT}/><wp:docPr id="${String(id)}" name="Chart ${String(id)}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${C}"><c:chart r:id="${rId}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const values = (cached: boolean): string =>
  `<c:val><c:numRef><c:f>Sheet!$B$2:$B$4</c:f>${
    cached
      ? `<c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>4.5</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt><c:pt idx="2"><c:v>3.5</c:v></c:pt></c:numCache>`
      : ""
  }</c:numRef></c:val>`;

const categories = (cached: boolean): string =>
  `<c:cat><c:strRef><c:f>Sheet!$A$2:$A$4</c:f>${
    cached
      ? `<c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>One</c:v></c:pt><c:pt idx="1"><c:v>Two</c:v></c:pt><c:pt idx="2"><c:v>Three</c:v></c:pt></c:strCache>`
      : ""
  }</c:strRef></c:cat>`;

const seriesXml = (cached: boolean): string =>
  `<c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="${BAR_HEX}"/></a:solidFill></c:spPr>${categories(cached)}${values(cached)}</c:ser>`;

const chartPart = (plot: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`;

const CHARTS = {
  cachedBar: chartPart(
    `<c:barChart><c:barDir val="col"/>${seriesXml(true)}</c:barChart>`,
  ),
  uncachedBar: chartPart(
    `<c:barChart><c:barDir val="col"/>${seriesXml(false)}</c:barChart>`,
  ),
  area: chartPart(`<c:areaChart>${seriesXml(true)}</c:areaChart>`),
} as const;

const NAMES = Object.keys(CHARTS) as (keyof typeof CHARTS)[];

const chartDocument = async (): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${NAMES.map(
      (_, index) =>
        `<Override PartName="/word/charts/chart${String(index)}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    ).join("")}</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${NAMES.map(
      (_, index) =>
        `<Relationship Id="rC${String(index)}" Type="${R}/chart" Target="charts/chart${String(index)}.xml"/>`,
    ).join("")}</Relationships>`,
  );
  NAMES.forEach((name, index) => {
    zip.set(`word/charts/chart${String(index)}.xml`, CHARTS[name]);
  });
  const body = NAMES.map(
    (name, index) =>
      paragraph(`BEFORE-${name}`) +
      chartParagraph(`rC${String(index)}`, index + 1) +
      paragraph(`AFTER-${name}`),
  ).join("");
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="31680"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

/** Vertical distance between the markers either side of one chart. */
const frameGap = async (page: Page, name: string): Promise<number> =>
  page
    .locator(".layout-page")
    .first()
    .evaluate((root, chart) => {
      const topOf = (text: string): number => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          if (walker.currentNode.textContent?.includes(text)) {
            const range = document.createRange();
            range.selectNodeContents(walker.currentNode);
            return range.getBoundingClientRect().top;
          }
        }
        throw new TypeError(`Missing marker ${text}`);
      };
      return topOf(`AFTER-${chart}`) - topOf(`BEFORE-${chart}`);
    }, name);

/** How many pixels of the page's screenshot have exactly this colour. */
const pixelsOfColour = async (
  page: Page,
  rgb: readonly [number, number, number],
): Promise<number> => {
  const shot = await page.locator(".layout-page").first().screenshot();
  return page.evaluate(
    async ({ png, colour }) => {
      const image = new Image();
      image.src = `data:image/png;base64,${png}`;
      await image.decode();
      const canvas = document.createElement("canvas");
      canvas.width = image.naturalWidth;
      canvas.height = image.naturalHeight;
      const context = canvas.getContext("2d");
      if (!context) throw new TypeError("No canvas context");
      context.drawImage(image, 0, 0);
      const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
      let count = 0;
      for (let index = 0; index < data.length; index += 4) {
        if (
          data[index] === colour[0] &&
          data[index + 1] === colour[1] &&
          data[index + 2] === colour[2]
        ) {
          count += 1;
        }
      }
      return count;
    },
    { png: shot.toString("base64"), colour: rgb },
  );
};

for (const { painter, query } of [
  { painter: "default", query: "?session=legacy" },
  {
    painter: "display-list",
    query: "?session=legacy&pageRenderer=display-list",
  },
]) {
  test(`charts paint and keep their frames on the ${painter} painter`, async ({
    page,
  }) => {
    await page.goto(`/${query}`);
    await page.locator("#file-input").setInputFiles({
      name: "charts.docx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      buffer: await chartDocument(),
    });
    await expect(page.locator(".layout-page").first()).toContainText(
      "AFTER-area",
    );

    for (const name of NAMES) {
      expect(
        await frameGap(page, name),
        `${name} keeps the height of its frame`,
      ).toBeGreaterThanOrEqual(CHART_HEIGHT_PX);
    }
    // Only the cached bar chart has marks in this colour; the series cannot
    // draw on an uncached or area chart, and those frames stay empty.
    await expect
      .poll(() => pixelsOfColour(page, BAR_RGB))
      .toBeGreaterThan(1000);
  });
}

test("a painted chart is an image the browser could decode", async ({
  page,
}) => {
  await page.goto("/?session=legacy");
  await page.locator("#file-input").setInputFiles({
    name: "charts.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await chartDocument(),
  });
  const images = page.locator(".layout-page img");
  await expect(images).toHaveCount(NAMES.length);
  await expect
    .poll(() =>
      images.evaluateAll((elements) =>
        elements.every(
          (element) =>
            element instanceof HTMLImageElement &&
            element.complete &&
            element.naturalWidth > 0 &&
            element.src.startsWith("data:image/svg+xml"),
        ),
      ),
    )
    .toBe(true);
});
