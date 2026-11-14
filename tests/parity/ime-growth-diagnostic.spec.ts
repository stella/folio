import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
  openEditor,
} from "./parity-fixture";

// Diagnostic: how the native and painted carets move apart while a CJK
// composition grows, and which fonts each side measures the text with.
for (const [engine, open] of [
  ["legacy", openEditor],
  ["canonical", openCanonicalEditor],
] as const) {
  forEachAdapter(
    `ime growth diagnostic ${engine}`,
    async (adapter, { page }) => {
      await open(page, adapter);
      await ensureLiveView(page);
      await page
        .locator(".paged-editor__hidden-pm > .ProseMirror")
        .first()
        .focus();
      await page
        .locator(".layout-page-content span[data-pm-start][data-pm-end]")
        .first()
        .click({ position: { x: 80, y: 8 } });
      await page.waitForTimeout(300);
      const read = () =>
        page.evaluate(() => {
          const selection = document.getSelection();
          const painted = document
            .querySelector<HTMLElement>("[data-folio-caret-rect]")
            ?.getBoundingClientRect();
          const host = document.querySelector<HTMLElement>(
            ".paged-editor__hidden-pm",
          );
          if (!selection || selection.rangeCount === 0) return "no-selection";
          const range = selection.getRangeAt(0);
          const native = range.getBoundingClientRect();
          const anchor = range.startContainer;
          const element =
            anchor instanceof Element ? anchor : anchor.parentElement;
          const style = element ? getComputedStyle(element) : null;
          return {
            native: [Math.round(native.left), Math.round(native.top)],
            painted: painted
              ? [Math.round(painted.left), Math.round(painted.top)]
              : null,
            transform: host?.style.transform ?? null,
            hiddenFont: style ? `${style.fontFamily} ${style.fontSize}` : null,
            text: (anchor.textContent ?? "").slice(0, 30),
          };
        });
      const log: unknown[] = [await read()];
      const cdp = await page.context().newCDPSession(page);
      const text = "あいうえおかきくけこさしすせそ";
      for (let length = 1; length <= text.length; length += 2) {
        await cdp.send("Input.imeSetComposition", {
          text: text.slice(0, length),
          selectionStart: length,
          selectionEnd: length,
        });
        await page.waitForTimeout(400);
        log.push({ length, ...((await read()) as object) });
      }
      const paintedFont = await page.evaluate(() => {
        const spans = [
          ...document.querySelectorAll<HTMLElement>(
            ".layout-page-content span",
          ),
        ].filter((span) => /[぀-ヿ]/u.test(span.textContent ?? ""));
        const span = spans.at(0);
        if (!span) return "no painted kana";
        const style = getComputedStyle(span);
        const canvas = document.createElement("canvas").getContext("2d");
        const widths: Record<string, number> = {};
        for (const family of [style.fontFamily, "serif", "sans-serif"]) {
          if (!canvas) break;
          canvas.font = `${style.fontSize} ${family}`;
          widths[family] = canvas.measureText("あいうえお").width;
        }
        return {
          font: `${style.fontFamily} ${style.fontSize}`,
          width: Math.round(span.getBoundingClientRect().width),
          text: span.textContent,
          widths,
        };
      });
      await cdp.send("Input.insertText", { text });
      await cdp.detach();
      expect(JSON.stringify({ engine, log, paintedFont })).toBe("");
    },
  );
}
