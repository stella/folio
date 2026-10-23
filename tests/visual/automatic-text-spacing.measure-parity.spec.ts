import { expect, test } from "@playwright/test";

import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

for (const renderer of ["legacy", "display-list"]) {
  test(`${renderer}: automatic advances preserve source-boundary Range, click and caret geometry`, async ({
    page,
  }) => {
    const source = createEmptyDocument({ initialText: "𠀀Aé\u0301漢1" });
    const first = source.package.document.content.at(0);
    if (first?.type !== "paragraph")
      throw new TypeError("Expected a paragraph.");
    first.content = [
      {
        type: "run",
        formatting: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial" },
          fontSize: 24,
        },
        content: [{ type: "text", text: "𠀀" }],
      },
      {
        type: "run",
        formatting: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial" },
          fontSize: 24,
          color: { rgb: "112233" },
        },
        content: [{ type: "text", text: "Aé\u0301漢1" }],
      },
    ];
    await page.goto(`/?file=sample.docx&pageRenderer=${renderer}`);
    await page.waitForSelector(".layout-page");
    await page.evaluate(
      (document) =>
        globalThis.__folioPlayground?.getEditorRef()?.loadDocument(document),
      source,
    );
    await expect(page.locator(".layout-page").first()).toContainText(
      "𠀀Aé\u0301漢1",
    );
    await page.evaluate(async () => {
      await document.fonts.ready;
    });
    const advances = await page.evaluate(() => {
      const host = document.querySelector<HTMLElement>(".layout-page");
      if (host === null) throw new TypeError("Missing page.");
      return [...host.querySelectorAll<HTMLElement>("span[style]")]
        .filter((span) => Number.parseFloat(span.style.letterSpacing) > 0)
        .map((span) => {
          const range = document.createRange();
          range.selectNodeContents(span);
          const computed = getComputedStyle(span);
          const context = document.createElement("canvas").getContext("2d");
          if (context === null) throw new TypeError("Missing canvas.");
          context.font = `${computed.fontStyle} ${computed.fontWeight} ${computed.fontSize} ${computed.fontFamily}`;
          context.fontKerning =
            computed.fontKerning === "none" ? "none" : "normal";
          return {
            text: span.textContent,
            delta:
              range.getBoundingClientRect().width -
              context.measureText(span.textContent ?? "").width,
            gap: Number.parseFloat(span.style.letterSpacing),
          };
        });
    });
    expect(advances.map((advance) => advance.text)).toEqual([
      "𠀀",
      "é\u0301",
      "漢",
    ]);
    for (const advance of advances)
      expect(Math.abs(advance.delta - advance.gap)).toBeLessThan(0.1);
    for (const offset of [2, 5, 6]) {
      const boundary = await page.evaluate((sourceOffset) => {
        const run = [
          ...document.querySelectorAll<HTMLElement>(
            ".layout-page span[data-pm-start][data-pm-end]",
          ),
        ].find(
          (span) =>
            Number(span.dataset["pmStart"]) <= sourceOffset + 1 &&
            Number(span.dataset["pmEnd"]) >= sourceOffset + 1,
        );
        if (run === undefined) throw new TypeError("Missing source run.");
        const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);
        let remaining = sourceOffset + 1 - Number(run.dataset["pmStart"]);
        let textNode = walker.nextNode();
        while (
          textNode !== null &&
          remaining > (textNode.nodeValue?.length ?? 0)
        ) {
          remaining -= textNode.nodeValue?.length ?? 0;
          textNode = walker.nextNode();
        }
        if (textNode === null) throw new TypeError("Missing source boundary.");
        const range = document.createRange();
        range.setStart(run, 0);
        range.setEnd(textNode, remaining);
        const rect = range.getBoundingClientRect();
        return { x: rect.right, y: rect.top + rect.height / 2 };
      }, offset);
      await page.mouse.click(boundary.x + 0.1, boundary.y);
      await expect
        .poll(() =>
          page.evaluate(
            () =>
              globalThis.__folioPlayground
                ?.getEditorRef()
                ?.getEditorRef()
                ?.getView()?.state.selection.from,
          ),
        )
        .toBe(offset + 1);
      const caret = await page.getByTestId("caret").first().boundingBox();
      expect(caret).not.toBeNull();
      expect(Math.abs((caret?.x ?? 0) - boundary.x)).toBeLessThan(1);
    }
    first.formatting = { ...first.formatting, indentRight: 10000 };
    first.content = [
      {
        type: "run",
        formatting: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial" },
          fontSize: 24,
        },
        content: [{ type: "text", text: "𠀀" }],
      },
      {
        type: "run",
        formatting: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial" },
          fontSize: 24,
          color: { rgb: "112233" },
        },
        content: [{ type: "text", text: "A" }],
      },
    ];
    await page.evaluate(
      (document) =>
        globalThis.__folioPlayground?.getEditorRef()?.loadDocument(document),
      source,
    );
    await expect(page.locator(".layout-page").first()).toContainText("𠀀A");
    await expect(page.locator(".layout-page").first()).not.toContainText("é");
    const wrapped = await page.evaluate(() => {
      const runs = [
        ...document.querySelectorAll<HTMLElement>(
          ".layout-page span[data-pm-start][data-pm-end]",
        ),
      ];
      const eastAsian = runs.find((run) => run.textContent === "𠀀");
      const latin = runs.find((run) => run.textContent === "A");
      return {
        before: eastAsian?.getBoundingClientRect().top,
        after: latin?.getBoundingClientRect().top,
        automaticSpans: [
          ...document.querySelectorAll<HTMLElement>(".layout-page span[style]"),
        ].filter((span) => Number.parseFloat(span.style.letterSpacing) > 0)
          .length,
      };
    });
    expect(wrapped.before).toBeDefined();
    expect(wrapped.after).toBeGreaterThan(wrapped.before ?? Infinity);
    expect(wrapped.automaticSpans).toBe(0);
  });
}
