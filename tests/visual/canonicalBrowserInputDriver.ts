import type { Page } from "@playwright/test";
import type { BrowserInputAction } from "./browserInputTrace";
import { driveBrowserIme } from "./browserImeDriver";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
export const driveCanonicalBrowserInput = async (page: Page, action: BrowserInputAction) => {
  switch (action.kind) {
    case "typing":
      await page.keyboard.insertText(action.text);
      return;
    case "enter":
      await page.keyboard.press("Enter");
      return;
    case "backspace":
      await page.keyboard.press("Backspace");
      return;
    case "delete":
      await page.keyboard.press("Delete");
      return;
    case "undo":
      await page.keyboard.press(`${MODIFIER}+z`);
      return;
    case "redo":
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      return;
    case "cut":
      await page.keyboard.press(`${MODIFIER}+x`);
      return;
    case "imeReplacement":
      await driveBrowserIme(page, action);
      return;
    case "historyBurst":
      for (const key of action.keys) {
        await page.keyboard.press(key === "undo" ? `${MODIFIER}+z` : `${MODIFIER}+Shift+z`);
      }
      return;
    case "pastePlain":
    case "pasteHtml":
    case "pasteWordHtml":
    case "pasteListHtml":
    case "pasteTable":
    case "pasteMultiBlock":
      await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
      await page.evaluate(async ({ plain, html }) => {
        const parts: Record<string, Blob> = {
          "text/plain": new Blob([plain], { type: "text/plain" }),
        };
        if (html) parts["text/html"] = new Blob([html], { type: "text/html" });
        await navigator.clipboard.write([new ClipboardItem(parts)]);
      }, action);
      await page.keyboard.press(`${MODIFIER}+v`);
      return;
    case "dragCellDelete":
    case "selectionDrag":
      throw new TypeError("Structural selections require a structural seed");
    default: {
      const unreachable: never = action;
      throw new TypeError(`Unknown browser input action: ${unreachable}`);
    }
  }
};
