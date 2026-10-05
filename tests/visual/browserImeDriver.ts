import type { Page } from "@playwright/test";
import type { BrowserInputAction } from "./browserInputTrace";
import type { CanonicalImeDriverCall } from "../parity/canonicalHistoryObservation";
import type {} from "../parity/canonicalBridge";
import { Result, TaggedError } from "better-result";

export class BrowserImeDriverCaptureError extends TaggedError("BrowserImeDriverCaptureError")<{
  message: string;
  cause: unknown;
  observationError: unknown;
  call: Pick<CanonicalImeDriverCall, "operation" | "text" | "before">;
}> {}

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
  let updates = 0;
  type ObserveCallOptions = {
    operation: CanonicalImeDriverCall["operation"];
    text: string;
    invoke: () => Promise<unknown>;
  };
  const observeCall = async ({ operation, text, invoke }: ObserveCallOptions) => {
    const before = await page.evaluate(() =>
      globalThis.__folioCanonicalImeDriverCalls === undefined
        ? null
        : (globalThis.__folioCanonical?.historyState("input") ?? null),
    );
    const delivered = await Result.tryPromise({ try: invoke, catch: (cause: unknown) => cause });
    const captured = await Result.tryPromise({
      try: () =>
        page.evaluate(
          (call) => {
            globalThis.__folioCanonicalImeDriverCalls?.push({
              ...call,
              after: globalThis.__folioCanonical?.historyState("input") ?? null,
            });
          },
          { operation, text, before },
        ),
      catch: (cause: unknown) => cause,
    });
    if (captured.isErr())
      throw new BrowserImeDriverCaptureError({
        message: "IME driver observation failed",
        cause: delivered.isErr() ? delivered.error : captured.error,
        observationError: captured.error,
        call: { operation, text, before },
      });
    if (delivered.isErr()) throw delivered.error;
  };
  try {
    await runBrowserImeLifecycle(
      {
        update: async (text) => {
          // CDP offsets count UTF-16 code units, including surrogate pairs.
          await observeCall({
            operation: updates++ === 0 ? "start" : "update",
            text,
            invoke: () =>
              cdp.send("Input.imeSetComposition", {
                text,
                selectionStart: text.length,
                selectionEnd: text.length,
              }),
          });
        },
        commit: async (text) => {
          await observeCall({
            operation: "commit",
            text,
            invoke: () => cdp.send("Input.insertText", { text }),
          });
        },
        cancel: async () => {
          await observeCall({
            operation: "cancel",
            text: "",
            invoke: () =>
              cdp.send("Input.imeSetComposition", {
                text: "",
                selectionStart: 0,
                selectionEnd: 0,
              }),
          });
        },
      },
      action,
    );
  } finally {
    await cdp.detach();
  }
};
