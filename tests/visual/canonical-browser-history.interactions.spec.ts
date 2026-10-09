import { evaluateCanonicalPage, observeCanonicalPageNavigation } from "./canonicalPageNavigation";
import { expect } from "@playwright/test";
import { test } from "./canonicalTimerProbe";
import fc from "fast-check";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import {
  checkCanonicalBrowserHistory,
  initializeCanonicalBrowserHistory,
} from "./canonicalBrowserHistoryOracle";
import { runBrowserImeLifecycle } from "./browserImeDriver";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  CANONICAL_BROWSER_SAVE_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";

test("canonical evaluation waits for a mid-sequence navigation to mount", async ({ page }) => {
  await page.goto("/?session=canonical");
  const source = await createDocx(createEmptyDocument({ initialText: "alpha" }));
  await initializeCanonicalBrowserHistory(page, [...new Uint8Array(source)]);
  expect(
    await evaluateCanonicalPage(page, () => page.evaluate(() => globalThis.__folioCanonicalReady)),
  ).toBe(true);

  let releaseResponse: () => void = () => {};
  const responseGate = new Promise<void>((resolve) => {
    releaseResponse = resolve;
  });
  let reachedNavigation: () => void = () => {};
  const navigationStarted = new Promise<void>((resolve) => {
    reachedNavigation = resolve;
  });
  await page.route("**/?session=canonical", async (route) => {
    reachedNavigation();
    await responseGate;
    await route.continue();
  });
  const reloaded = page.reload({ waitUntil: "load" });
  await navigationStarted;
  const navigation = observeCanonicalPageNavigation(page);
  let callbackInvoked = false;
  const evaluated = evaluateCanonicalPage(page, () => {
    callbackInvoked = true;
    return page.evaluate(() => ({
      ready: globalThis.__folioCanonicalReady,
      documentState: document.readyState,
      // A reload mounts a new bridge, before the next fixture load.
      composing: globalThis.__folioCanonical?.nativeComposing(),
    }));
  });
  try {
    // The old document is still mounted and complete while the response is
    // held; waiting only for its load state would invoke the callback here.
    expect(navigation.current.status).toBe("loading");
    if (navigation.current.status !== "loading")
      throw new TypeError("Held reload must retain its navigation owner");
    // The helper installs this barrier synchronously before its first await.
    // Removing the pending-navigation wait leaves it null, so the regression
    // fails without depending on transport timing or a fixed delay.
    expect(navigation.current.load).not.toBeNull();
    expect(callbackInvoked).toBe(false);
  } finally {
    releaseResponse();
  }
  await reloaded;
  expect(await evaluated).toEqual({ ready: true, documentState: "complete", composing: null });
  expect(callbackInvoked).toBe(true);
  await initializeCanonicalBrowserHistory(page, [...new Uint8Array(source)]);
  expect(
    await checkCanonicalBrowserHistory({
      page,
      source: [...new Uint8Array(source)],
      actions: [],
      missing: createMissingOpBurndown(),
    }),
  ).toBe(0);
});

test("canonical evaluation reports a failed mid-sequence navigation", async ({ page }) => {
  await page.goto("/?session=canonical");
  await evaluateCanonicalPage(page, () => page.evaluate(() => globalThis.__folioCanonicalReady));
  let releaseResponse: () => void = () => {};
  const responseGate = new Promise<void>((resolve) => {
    releaseResponse = resolve;
  });
  let reachedNavigation: () => void = () => {};
  const navigationStarted = new Promise<void>((resolve) => {
    reachedNavigation = resolve;
  });
  await page.route("**/?session=canonical", async (route) => {
    reachedNavigation();
    await responseGate;
    await route.abort("failed");
  });
  const reloaded = page.reload().then(
    () => "loaded",
    () => "failed",
  );
  await navigationStarted;
  let callbackInvoked = false;
  const evaluated = evaluateCanonicalPage(page, () => {
    callbackInvoked = true;
    return page.evaluate(() => globalThis.__folioCanonicalReady);
  });
  const failure = expect(evaluated).rejects.toThrow("Canonical playground navigation failed:");
  releaseResponse();
  await failure;
  expect(await reloaded).toBe("failed");
  expect(callbackInvoked).toBe(false);
});

test("refused native IME commit ends composition before driver completion", async ({ page }) => {
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  const source = await createDocx(createEmptyDocument({ initialText: "alpha" }));
  expect(
    await evaluateCanonicalPage(page, () =>
      page.evaluate(
        (bytes) => globalThis.__folioCanonical?.load(bytes),
        [...new Uint8Array(source)],
      ),
    ),
  ).toBe(true);
  expect(
    await evaluateCanonicalPage(page, () =>
      page.evaluate(() => globalThis.__folioCanonical?.select(1, 6)),
    ),
  ).toBe(true);
  await evaluateCanonicalPage(page, () =>
    page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
      const editor = document.querySelector<HTMLElement>(".ProseMirror");
      if (!editor) throw new TypeError("Canonical editor unavailable");
      editor.addEventListener("beforeinput", (event) => {
        if (event.inputType === "insertText" || event.inputType === "insertReplacementText")
          editor.dataset["finalPrevented"] = String(event.defaultPrevented);
      });
      editor.addEventListener("compositionend", (event) => {
        if (event.isTrusted) editor.dataset["nativeEndDelivered"] = "true";
      });
    }),
  );
  const cdp = await page.context().newCDPSession(page);
  let finished = false;
  try {
    await runBrowserImeLifecycle(
      {
        update: async (text) => {
          await cdp.send("Input.imeSetComposition", {
            text,
            selectionStart: text.length,
            selectionEnd: text.length,
          });
          // Multiple differently marked text runs avoid PM's single-text insertion
          // shortcut, which intentionally ignores parsed marks during native typing.
          await page.locator(".ProseMirror").evaluate((editor) => {
            const paragraph = editor.querySelector("p");
            if (!paragraph) throw new TypeError("Composition paragraph unavailable");
            editor.dispatchEvent(
              new InputEvent("beforeinput", {
                bubbles: true,
                inputType: "insertCompositionText",
                data: "契",
                isComposing: true,
              }),
            );
            const marked = document.createElement("strong");
            marked.textContent = "契約";
            paragraph.replaceChildren(document.createTextNode("甲"), marked);
            editor.dispatchEvent(
              new InputEvent("input", {
                bubbles: true,
                inputType: "insertCompositionText",
                data: "契",
                isComposing: true,
              }),
            );
          });
          await expect
            .poll(
              () =>
                evaluateCanonicalPage(page, () =>
                  page.evaluate(() =>
                    globalThis.__folioCanonicalFuzzErrors?.some((error) =>
                      error.message.includes("Composition changed unsupported content"),
                    ),
                  ),
                ),
              { message: "marked native composition must be refused before the final commit" },
            )
            .toBe(true);
          // A full snapshot is unavailable while canonical composition remains active.
          expect(
            await evaluateCanonicalPage(page, () =>
              page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
            ),
          ).toBe(true);
          expect(await page.locator(".ProseMirror").evaluate((editor) => editor.textContent)).toBe(
            "alpha",
          );
        },
        commit: async (text) => {
          await cdp.send("Input.insertText", { text });
          expect(finished).toBe(false);
          expect(await page.locator(".ProseMirror").getAttribute("data-final-prevented")).toBe(
            "true",
          );
          expect(await page.locator(".ProseMirror").getAttribute("data-native-end-delivered")).toBe(
            null,
          );
          // This assertion precedes driver completion: removing product composition exit makes it fail.
          expect(
            await evaluateCanonicalPage(page, () =>
              page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
            ),
          ).toBe(false);
          const current = await evaluateCanonicalPage(page, () =>
            page.evaluate(() => globalThis.__folioCanonical?.snapshot()),
          );
          expect(current?.composing).toBe(false);
          expect(current?.text).toBe("alpha");
          expect(current?.canUndo).toBe(false);
          expect(current?.projectionMatchesCanonical).toBe(true);
        },
        cancel: async () => {
          await cdp.send("Input.imeSetComposition", {
            text: "",
            selectionStart: 0,
            selectionEnd: 0,
          });
        },
        finish: async () => {
          finished = true;
          await page.locator(".ProseMirror").evaluate((editor) => {
            if (!(editor instanceof HTMLElement)) throw new TypeError("Editor unavailable");
            editor.blur();
            editor.focus();
          });
        },
      },
      { kind: "imeReplacement", updates: ["契"], completion: "commit" },
    );
    expect(finished).toBe(true);
    const final = await evaluateCanonicalPage(page, () =>
      page.evaluate(() => globalThis.__folioCanonical?.snapshot()),
    );
    expect(final?.text).toBe("alpha");
    expect(final?.composing).toBe(false);
    expect(final?.canUndo).toBe(false);
    expect(final?.projectionMatchesCanonical).toBe(true);
  } finally {
    if (!page.isClosed()) await cdp.detach();
  }
});

test("canonical history checks a lazy first view and rejects composition before reload", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  expect(
    await evaluateCanonicalPage(page, () =>
      page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
    ),
  ).toBe(null);
  const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
  await initializeCanonicalBrowserHistory(page, [...new Uint8Array(source)]);
  const options = {
    page,
    source: [...new Uint8Array(source)],
    actions: [],
    missing: createMissingOpBurndown(),
  };
  expect(await checkCanonicalBrowserHistory(options)).toBe(0);
  expect(
    await evaluateCanonicalPage(page, () =>
      page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
    ),
  ).toBe(false);
  await page.locator(".ProseMirror").evaluate((element) => {
    element.dispatchEvent(new Event("compositionstart", { bubbles: true }));
  });
  await expect(checkCanonicalBrowserHistory(options)).rejects.toThrow(
    "case must start outside native composition",
  );
  // The failed invariant must not load/reset the owner and hide the leak.
  expect(
    await evaluateCanonicalPage(page, () =>
      page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
    ),
  ).toBe(true);
});

for (const { seed, path, kinds } of [
  ...CANONICAL_BROWSER_HISTORY_REPLAYS,
  ...CANONICAL_BROWSER_SAVE_REPLAYS,
]) {
  test(`canonical history replay ${seed} ${path}`, async ({ page }) => {
    if (seed === 197 && path === "1") test.setTimeout(120_000);
    const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
    expect(traces).toHaveLength(1);
    const actions = traces.at(0);
    if (actions === undefined) throw new TypeError("Missing canonical regression trace");
    expect(actions.length).toBeGreaterThan(0);
    expect(actions.map(({ kind }) => kind)).toEqual(kinds);
    if (seed === 431 && path === "8")
      expect(actions).toEqual([
        { kind: "imeReplacement", updates: ["shall", "café 東京 é"], completion: "cancel" },
      ]);
    await page.goto("/?session=canonical");
    await page.waitForSelector(".layout-page");
    await evaluateCanonicalPage(page, () =>
      page.evaluate(() => {
        globalThis.__folioCanonicalFuzzErrors = [];
      }),
    );
    const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
    await initializeCanonicalBrowserHistory(page, [...new Uint8Array(source)]);
    // Repeat identical package input to exercise adoption of a fresh owner.
    for (let load = 0; load < 2; load++) {
      const applied = await checkCanonicalBrowserHistory({
        page,
        source: [...new Uint8Array(source)],
        actions,
        missing: createMissingOpBurndown(),
      });
      if (
        actions.every(
          (action) => action.kind === "imeReplacement" && action.completion === "cancel",
        )
      )
        expect(applied).toBe(0);
      else expect(applied).toBeGreaterThan(0);
    }
  });
}
