import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const para = (text: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const lines = (count: number) =>
  Array.from(
    { length: count },
    (_, index) => `G${String(index).padStart(3, "0")}`,
  );

type Note = {
  count: number;
  /** How the note holds its lines. */
  noteAs?: "paragraphs" | "spaced paragraphs" | "a table";
  bodyLines: number;
};

const noteBody = ({ count, noteAs = "paragraphs" }: Note) => {
  if (noteAs === "a table") {
    const rows = lines(count)
      .map(
        (text) =>
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc></w:tr>`,
      )
      .join("");
    return `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl><w:p/>`;
  }
  if (noteAs === "spaced paragraphs") {
    return lines(count)
      .map(
        (text) =>
          `<w:p><w:pPr><w:spacing w:before="120" w:after="200"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`,
      )
      .join("");
  }
  return lines(count).map(para).join("");
};

const documentWith = async (note: Note) => {
  const { bodyLines } = note;
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdP" Type="${R}/footer" Target="footer1.xml"/></Relationships>`,
  );
  zip.set(
    "word/footer1.xml",
    `${HEAD}<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">Sheet </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
  );
  const body = Array.from({ length: bodyLines }, (_, index) =>
    para(`Body ${String(index)}`),
  ).join("");
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Reference here</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>${body}${para("TAIL")}<w:sectPr xmlns:r="${R}"><w:footerReference w:type="default" r:id="rIdP"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/footnotes.xml",
    `${HEAD}<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${noteBody(note)}</w:footnote></w:footnotes>`,
  );
  return Buffer.from(await zip.generate());
};

type Seen = { text: string; whole: boolean; visible: boolean; page: number };

const paintedLines = (page: Page): Promise<Seen[]> =>
  page.evaluate(() => {
    const seen: Seen[] = [];
    const pages = [...document.querySelectorAll<HTMLElement>(".layout-page")];
    for (const [index, pageElement] of pages.entries()) {
      const pageBox = pageElement.getBoundingClientRect();
      const walker = document.createTreeWalker(
        pageElement,
        NodeFilter.SHOW_TEXT,
      );
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = (node.textContent ?? "").trim();
        if (!/^G\d{3}$/u.test(text)) continue;
        const range = document.createRange();
        range.selectNodeContents(node);
        const box = range.getBoundingClientRect();
        let top = Math.max(box.top, pageBox.top);
        let bottom = Math.min(box.bottom, pageBox.bottom);
        for (
          let owner = node.parentElement;
          owner && owner !== pageElement;
          owner = owner.parentElement
        ) {
          if (getComputedStyle(owner).overflowY === "visible") continue;
          const clip = owner.getBoundingClientRect();
          top = Math.max(top, clip.top);
          bottom = Math.min(bottom, clip.bottom);
        }
        const shown = bottom - top;
        seen.push({
          text,
          page: index,
          visible: shown > 2,
          // A text box runs past its line box, so a clip at a line bottom can
          // shave it by a pixel or two with every glyph showing.
          whole: shown >= box.height - 2,
        });
      }
    }
    return seen;
  });

/** Body lines painted lower than a note line that shows on the same page. */
const bodyLinesInsideNotes = (page: Page): Promise<string[]> =>
  page.evaluate(() => {
    const inside: string[] = [];
    for (const pageElement of document.querySelectorAll<HTMLElement>(
      ".layout-page",
    )) {
      let noteTop = Number.POSITIVE_INFINITY;
      let lowestBody = 0;
      let lowestText = "";
      const pageBox = pageElement.getBoundingClientRect();
      const walker = document.createTreeWalker(
        pageElement,
        NodeFilter.SHOW_TEXT,
      );
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = (node.textContent ?? "").trim();
        const range = document.createRange();
        range.selectNodeContents(node);
        const box = range.getBoundingClientRect();
        if (box.bottom < pageBox.top || box.top > pageBox.bottom) continue;
        if (/^G\d{3}$/u.test(text)) {
          let clip = box.top;
          for (
            let owner = node.parentElement;
            owner && owner !== pageElement;
            owner = owner.parentElement
          ) {
            if (getComputedStyle(owner).overflowY !== "visible") {
              clip = Math.max(clip, owner.getBoundingClientRect().top);
            }
          }
          noteTop = Math.min(noteTop, clip);
        } else if (/^(Body \d+|TAIL|Reference here)/u.test(text)) {
          if (box.bottom > lowestBody) {
            lowestBody = box.bottom;
            lowestText = text;
          }
        }
      }
      if (lowestBody > noteTop + 1) inside.push(lowestText);
    }
    return inside;
  });

for (const [painter, url] of [
  ["the default painter", "/?session=canonical"],
  ["the display-list painter", "/?session=canonical&pageRenderer=display-list"],
] as const) {
  for (const [shape, note] of [
    ["the next page", { count: 150, bodyLines: 0 }],
    ["several pages", { count: 300, bodyLines: 40 }],
    [
      "pages, with paragraph spacing",
      { count: 80, noteAs: "spaced paragraphs", bodyLines: 5 },
    ],
    ["pages, as a table", { count: 120, noteAs: "a table", bodyLines: 5 }],
  ] satisfies [string, Note][]) {
    const { count } = note;
    test(`a long footnote continues across ${shape}, on ${painter}`, async ({
      page,
    }) => {
      await page.goto(url);
      await page.locator("#file-input").setInputFiles({
        name: "note.docx",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        buffer: await documentWith(note),
      });
      await expect
        .poll(
          async () =>
            (await page.locator(".layout-page").allTextContents()).join(""),
          { timeout: 30_000 },
        )
        .toContain("TAIL");
      await page.waitForTimeout(1000);
      const painted = await paintedLines(page);
      const whole = painted.filter((line) => line.whole);
      expect(
        lines(count).filter(
          (text) => !whole.some((line) => line.text === text),
        ),
        "every line of the note shows whole",
      ).toEqual([]);
      expect(
        whole
          .map(({ text }) => text)
          .filter((text, index, all) => all.indexOf(text) !== index),
        "no line of the note shows twice",
      ).toEqual([]);
      expect(
        painted.filter(({ visible, whole: all }) => visible && !all),
        "no line of the note is cut",
      ).toEqual([]);
      expect(
        await bodyLinesInsideNotes(page),
        "no body line stands in a note area",
      ).toEqual([]);
      const sheets = await page.locator(".layout-page").allTextContents();
      expect(
        sheets.flatMap((text) => /Sheet (\d+)/u.exec(text)?.[1] ?? []),
        "every page keeps its own number",
      ).toEqual(sheets.map((_, index) => String(index + 1)));
    });
  }
}
