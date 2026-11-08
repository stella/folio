import type { Page } from "@playwright/test";

import type { BrowserInputAction } from "./browserInputTrace";
import type {} from "./browserTestBridge";

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
    if (
      !(editor instanceof HTMLElement) ||
      !editor.classList.contains("ProseMirror")
    )
      throw new TypeError("Native IME requires a focused editor");
    let phase: "active" | "ended" = "active";
    const controller = new AbortController();
    const started = performance.now();
    const diag: string[] = [];
    const describe = (node: Node | null) => {
      if (!node) return "null";
      const element = node instanceof Element ? node : node.parentElement;
      const owner = element?.closest(".ProseMirror");
      return `${node.nodeName}${owner === editor ? "@editor" : owner ? "@other" : ""}${node.isConnected ? "" : "!detached"}:${JSON.stringify((node.textContent ?? "").slice(0, 40))}`;
    };
    const selectionText = () => {
      const selection = document.getSelection();
      return selection
        ? `${describe(selection.anchorNode)}+${selection.anchorOffset}`
        : "none";
    };
    const geometry = () => {
      const host = editor.parentElement;
      const rect = editor.getBoundingClientRect();
      const selection = document.getSelection();
      const caret =
        selection && selection.rangeCount > 0
          ? selection.getRangeAt(0).getBoundingClientRect()
          : null;
      return `geom[host=${host?.style.transform ?? ""} editor=${Math.round(rect.left)},${Math.round(rect.top)},${Math.round(rect.width)}x${Math.round(rect.height)} caret=${caret ? `${Math.round(caret.left)},${Math.round(caret.top)},${Math.round(caret.height)}` : "none"}]`;
    };
    const note = (entry: string) => {
      const view = globalThis.__folioPlayground
        ?.getEditorRef()
        ?.getEditorRef()
        ?.getView();
      diag.push(
        `${Math.round(performance.now() - started)} ${entry} ${geometry()} pm[composing=${view?.composing} same=${view?.dom === editor} sel=${view?.state.selection.from}-${view?.state.selection.to}]`,
      );
    };
    for (const type of [
      "compositionstart",
      "compositionupdate",
      "compositionend",
      "beforeinput",
      "input",
      "focusin",
      "focusout",
    ])
      document.addEventListener(
        type,
        (event) => {
          const data =
            event instanceof CompositionEvent || event instanceof InputEvent
              ? JSON.stringify(event.data)
              : "";
          const inputType = event instanceof InputEvent ? event.inputType : "";
          note(
            `${type} ${inputType} ${data} trusted=${event.isTrusted} target=${describe(event.target as Node)} active=${describe(document.activeElement)} sel=${selectionText()}`,
          );
        },
        { capture: true, signal: controller.signal },
      );
    const observer = new MutationObserver((records) => {
      for (const record of records)
        note(
          `mutation ${record.type} target=${describe(record.target)} +${record.addedNodes.length} -${record.removedNodes.length} removed=${[...record.removedNodes].map(describe).join("|")}`,
        );
    });
    observer.observe(editor, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ["class", "contenteditable"],
    });
    controller.signal.addEventListener("abort", () => observer.disconnect());
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
      diag,
      note,
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
          await lifecycle.evaluate(
            ({ note }, value) => note(`STEP update ${JSON.stringify(value)}`),
            text,
          );
          // CDP offsets count UTF-16 code units, including surrogate pairs.
          await cdp.send("Input.imeSetComposition", {
            text,
            selectionStart: text.length,
            selectionEnd: text.length,
          });
        },
        commit: async (text) => {
          await lifecycle.evaluate(
            ({ note }, value) => note(`STEP commit ${JSON.stringify(value)}`),
            text,
          );
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
    const diag = await lifecycle.evaluate(({ diag }) => diag.slice());
    console.log(
      `IME-DIAG ${JSON.stringify(action.updates)}\n${diag.join("\n")}`,
    );
    await lifecycle.evaluate(({ controller }) => controller.abort());
    await lifecycle.dispose();
    await cdp.detach();
  }
};
