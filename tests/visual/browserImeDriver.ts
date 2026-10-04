import type { Page } from "@playwright/test";
import type { BrowserInputAction } from "./browserInputTrace";

type BrowserImeAction = Extract<BrowserInputAction, { kind: "imeReplacement" }>;
type ImeLifecycleDriver = {
  update: (text: string) => Promise<void>;
  commit: (text: string) => Promise<void>;
  cancel: () => Promise<void>;
};
export const runBrowserImeLifecycle = async (
  driver: ImeLifecycleDriver,
  action: BrowserImeAction,
) => {
  const last = action.updates.at(-1);
  if (last === undefined) throw new TypeError("IME lifecycle has no updates");
  for (const text of action.updates) await driver.update(text);
  switch (action.completion) {
    case "commit":
      await driver.commit(last);
      return;
    case "cancel":
      await driver.cancel();
      return;
    default: {
      const unexpected: never = action.completion;
      throw new TypeError(`Unknown IME completion: ${unexpected}`);
    }
  }
};
export const driveBrowserIme = async (page: Page, action: BrowserImeAction) => {
  const cdp = await page.context().newCDPSession(page);
  try {
    await runBrowserImeLifecycle(
      {
        update: async (text) => {
          // CDP offsets count UTF-16 code units, including surrogate pairs.
          await cdp.send("Input.imeSetComposition", {
            text,
            selectionStart: text.length,
            selectionEnd: text.length,
          });
        },
        commit: async (text) => {
          await cdp.send("Input.insertText", { text });
        },
        cancel: async () => {
          await cdp.send("Input.imeSetComposition", {
            text: "",
            selectionStart: 0,
            selectionEnd: 0,
          });
        },
      },
      action,
    );
  } finally {
    await cdp.detach();
  }
};
