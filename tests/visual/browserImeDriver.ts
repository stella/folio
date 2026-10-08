import type { Page } from "@playwright/test";
import type { BrowserInputAction } from "./browserInputTrace";

type BrowserImeAction = Extract<BrowserInputAction, { kind: "imeReplacement" }>;
type ImeLifecycleDriver = {
  update: (text: string) => Promise<void>;
  commit: (text: string) => Promise<void>;
  cancel: () => Promise<void>;
  finish: (completion: BrowserImeAction["completion"]) => Promise<void>;
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
      break;
    case "cancel":
      await driver.cancel();
      break;
    default: {
      const unexpected: never = action.completion;
      throw new TypeError(`Unknown IME completion: ${unexpected}`);
    }
  }
  // A refused native edit can leave the browser IME owner active. Finish its
  // focus lifecycle without issuing a second edit against restored canonical DOM.
  await driver.finish(action.completion);
};
export const driveBrowserIme = async (page: Page, action: BrowserImeAction) => {
  const lifecycle = await page.evaluateHandle(() => {
    const editor = document.activeElement;
    if (!(editor instanceof HTMLElement) || !editor.classList.contains("ProseMirror"))
      throw new TypeError("Native IME requires a focused editor");
    let phase: "active" | "ended" = "active";
    const controller = new AbortController();
    editor.addEventListener(
      "compositionend",
      (event) => {
        // Product recovery ends PM through a plain Event. Only the browser
        // CompositionEvent confirms that its own IME lifecycle has ended.
        if (event instanceof CompositionEvent) phase = "ended";
      },
      { signal: controller.signal },
    );
    return {
      editor,
      controller,
      get phase() {
        return phase;
      },
    };
  });
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
        finish: async (completion) => {
          if (await lifecycle.evaluate(({ phase }) => phase === "ended")) return;
          // An explicit cancellation must discard a pending projection before
          // blur recovery can commit it when Chromium omitted the end event.
          if (completion === "cancel") await page.keyboard.press("Escape");
          await lifecycle.evaluate(({ editor, controller, phase }) => {
            controller.abort();
            if (phase === "ended") return;
            editor.blur();
            editor.focus();
          });
        },
      },
      action,
    );
  } finally {
    await lifecycle.evaluate(({ controller }) => controller.abort());
    await lifecycle.dispose();
    await cdp.detach();
  }
};
