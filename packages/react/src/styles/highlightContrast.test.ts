import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";

const BACKGROUND_TEXT_SELECTOR =
  '[style*="--doc-run-background-text-color"].docx-run-background-text';
const EXPLICIT_COLOR_SELECTOR =
  '.dark .layout-page [style*="--doc-run-color"]:not(.docx-insertion):not(.docx-deletion)';
const AUTHORED_BACKGROUND_SELECTOR =
  '.dark .layout-page [style*="--doc-authored-background-color"]';

describe("dark-mode authored run backgrounds", () => {
  test("inverts authored backgrounds and leaves their text on the document color path", () => {
    const css = readFileSync(new URL("editor.css", import.meta.url), "utf-8");
    const explicitColorRule = css.indexOf(EXPLICIT_COLOR_SELECTOR);
    const authoredBackgroundRule = css.indexOf(AUTHORED_BACKGROUND_SELECTOR);

    expect(explicitColorRule).toBeGreaterThanOrEqual(0);
    expect(authoredBackgroundRule).toBeGreaterThan(explicitColorRule);
    const ruleStart = css.lastIndexOf(".dark", authoredBackgroundRule);
    expect(css.slice(ruleStart, css.indexOf("}", authoredBackgroundRule) + 1)).toContain(
      "from var(--doc-authored-background-color) clamp(0, 1.22 - 0.95 * l, 1)",
    );
    expect(css).not.toContain(BACKGROUND_TEXT_SELECTOR);
  });
});
