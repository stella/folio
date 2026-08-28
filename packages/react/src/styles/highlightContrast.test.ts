import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";

const BACKGROUND_TEXT_SELECTOR =
  '[style*="--doc-run-background-text-color"].docx-run-background-text';
const EXPLICIT_COLOR_SELECTOR =
  '.dark .layout-page [style*="--doc-run-color"]:not(.docx-insertion):not(.docx-deletion)';

describe("dark-mode authored run backgrounds", () => {
  test("the local background foreground wins after ordinary run-color inversion", () => {
    const css = readFileSync(new URL("editor.css", import.meta.url), "utf-8");
    const explicitColorRule = css.indexOf(EXPLICIT_COLOR_SELECTOR);
    const backgroundTextRule = css.indexOf(BACKGROUND_TEXT_SELECTOR);

    expect(explicitColorRule).toBeGreaterThanOrEqual(0);
    expect(backgroundTextRule).toBeGreaterThan(explicitColorRule);
    const ruleStart = css.lastIndexOf(".dark", backgroundTextRule);
    expect(css.slice(ruleStart, css.indexOf("}", backgroundTextRule) + 1)).toContain(
      "color: var(--doc-run-background-text-color) !important",
    );
  });
});
