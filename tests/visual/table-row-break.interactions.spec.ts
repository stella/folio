import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const lines = (prefix: string, count: number, lineTwips: number): string =>
  Array.from(
    { length: count },
    (_, index) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${String(lineTwips)}" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>${prefix}${String(index).padStart(3, "0")}</w:t></w:r></w:p>`,
  ).join("");

/**
 * One row taller than a page. 263 and 347 twips are prime, so its cells' lines
 * share no bottom before 91,261 twips, far past the bottom of any page.
 */
const documentWith = async (left: number, right: number): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const cell = (content: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(lines("A", left, 263))}${cell(lines("B", right, 347))}</w:tr></w:tbl><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

type Seen = { text: string; page: number; whole: boolean; visible: boolean };

/** Every line text, on the page it is painted on, and how much of it shows. */
const paintedLines = (page: Page): Promise<Seen[]> =>
  page.evaluate(() => {
    const seen: Seen[] = [];
    const pages = [...document.querySelectorAll<HTMLElement>(".layout-page")];
    for (const [index, pageElement] of pages.entries()) {
      const walker = document.createTreeWalker(
        pageElement,
        NodeFilter.SHOW_TEXT,
      );
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = (node.textContent ?? "").trim();
        if (!/^[AB]\d{3}$/u.test(text)) continue;
        const range = document.createRange();
        range.selectNodeContents(node);
        const box = range.getBoundingClientRect();
        const pageBox = pageElement.getBoundingClientRect();
        let top = Math.max(box.top, pageBox.top);
        let bottom = Math.min(box.bottom, pageBox.bottom);
        // What clips a line within its page. The scroller outside the page
        // only hides what is out of view.
        for (
          let owner = node.parentElement;
          owner && owner !== pageElement;
          owner = owner.parentElement
        ) {
          const style = getComputedStyle(owner);
          const clip = owner.getBoundingClientRect();
          const inset = /^inset\(([^)]*)\)/u.exec(style.clipPath)?.[1];
          if (inset) {
            const [above = "0", , below = above] = inset.split(" ");
            top = Math.max(top, clip.top + Number.parseFloat(above));
            bottom = Math.min(bottom, clip.bottom - Number.parseFloat(below));
          }
          if (style.overflowY !== "visible") {
            top = Math.max(top, clip.top);
            bottom = Math.min(bottom, clip.bottom);
          }
        }
        const shown = bottom - top;
        seen.push({
          text,
          page: index,
          visible: shown > 0.5,
          whole: shown >= box.height - 0.5,
        });
      }
    }
    return seen;
  });

for (const [painter, url] of [
  ["the default painter", "/?session=canonical"],
  ["the display-list painter", "/?session=canonical&pageRenderer=display-list"],
] as const) {
  test(`a row whose cells share no line boundary shows every line once, on ${painter}`, async ({
    page,
  }) => {
    const left = 200;
    const right = 150;
    await page.goto(url);
    await page.locator("#file-input").setInputFiles({
      name: "row-break.docx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      buffer: await documentWith(left, right),
    });
    await expect(page.locator(".layout-page").last()).toContainText("TAIL", {
      timeout: 30_000,
    });
    const painted = await paintedLines(page);
    expect(new Set(painted.map(({ page: at }) => at)).size).toBeGreaterThan(1);
    const expected = [
      ...Array.from(
        { length: left },
        (_, index) => `A${String(index).padStart(3, "0")}`,
      ),
      ...Array.from(
        { length: right },
        (_, index) => `B${String(index).padStart(3, "0")}`,
      ),
    ];
    const shownWhole = painted.filter(({ whole }) => whole);
    expect(
      expected.filter((text) => !shownWhole.some((line) => line.text === text)),
      "every line shows whole somewhere",
    ).toEqual([]);
    expect(
      shownWhole
        .map(({ text }) => text)
        .filter((text, index, all) => all.indexOf(text) !== index),
      "no line shows twice",
    ).toEqual([]);
    expect(
      painted.filter(({ visible, whole }) => visible && !whole),
      "no line is cut",
    ).toEqual([]);
  });
}
