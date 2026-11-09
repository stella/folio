import type { Page } from "@playwright/test";

import type { BrowserInputAction } from "./browserInputTrace";

type BrowserImeAction = Extract<BrowserInputAction, { kind: "imeReplacement" }>;
/**
 * The composing paragraph's text after one update, and its text before and
 * after the selection the composition started from.
 */
export type ProvisionalReading = {
  update: string;
  before: string;
  after: string;
  actual: string | null;
};
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
/**
 * Drives one composition through Chromium's input method. With
 * `readProvisional`, it reads the composing paragraph after each update: the
 * text in progress belongs where the caret was, between the paragraph's text
 * before and after the selection it started from.
 */
export const driveBrowserIme = async (
  page: Page,
  action: BrowserImeAction,
  { readProvisional = false }: { readProvisional?: boolean } = {},
): Promise<readonly ProvisionalReading[]> => {
  const lifecycle = await page.evaluateHandle(() => {
    const editor = document.activeElement;
    if (
      !(editor instanceof HTMLElement) ||
      !editor.classList.contains("ProseMirror")
    )
      throw new TypeError("Native IME requires a focused editor");
    let phase: "active" | "ended" = "active";
    const controller = new AbortController();
    const paragraphOf = (node: Node) =>
      (node instanceof Element ? node : node.parentElement)?.closest("p") ??
      null;
    const range =
      (document.getSelection()?.rangeCount ?? 0) > 0
        ? document.getSelection()?.getRangeAt(0)
        : undefined;
    const started = range ? paragraphOf(range.startContainer) : null;
    // Only a composition inside one paragraph has a text to expect there.
    const around =
      range &&
      started &&
      editor.contains(started) &&
      paragraphOf(range.endContainer) === started
        ? (() => {
            const before = document.createRange();
            before.selectNodeContents(started);
            before.setEnd(range.startContainer, range.startOffset);
            const after = document.createRange();
            after.selectNodeContents(started);
            after.setStart(range.endContainer, range.endOffset);
            return {
              paragraph: started,
              id: started.getAttribute("data-para-id"),
              before: before.toString(),
              after: after.toString(),
            };
          })()
        : null;
    /** The composing paragraph's text, even when the editor redrew it. */
    const provisional = () => {
      if (around === null) return null;
      const redrawn =
        around.id === null
          ? null
          : editor.querySelector(`p[data-para-id="${CSS.escape(around.id)}"]`);
      const paragraph = around.paragraph.isConnected
        ? around.paragraph
        : redrawn;
      return {
        before: around.before,
        after: around.after,
        actual: paragraph?.textContent ?? null,
      };
    };
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
      provisional,
      get phase() {
        return phase;
      },
    };
  });
  const readings: ProvisionalReading[] = [];
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
          if (!readProvisional) return;
          // Read once the editor has handled the update and painted it.
          const read = await lifecycle.evaluate(
            ({ provisional }) =>
              new Promise<ReturnType<typeof provisional>>((resolve) =>
                requestAnimationFrame(() =>
                  setTimeout(() => resolve(provisional()), 0),
                ),
              ),
          );
          if (read !== null) readings.push({ update: text, ...read });
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
          if (await lifecycle.evaluate(({ phase }) => phase === "ended"))
            return;
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
  return readings;
};
