import { expect, test } from "@playwright/test";
import fc from "fast-check";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { checkCanonicalBrowserHistory } from "./canonicalBrowserHistoryOracle";
import { runBrowserImeLifecycle } from "./browserImeDriver";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";

test("refused native IME commit ends composition before driver cancellation", async ({ page }) => {
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  const source = await createDocx(createEmptyDocument({ initialText: "alpha" }));
  expect(
    await page.evaluate(
      (bytes) => globalThis.__folioCanonical?.load(bytes),
      [...new Uint8Array(source)],
    ),
  ).toBe(true);
  expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
  await page.evaluate(() => {
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
  });
  const cdp = await page.context().newCDPSession(page);
  let cancelled = false;
  try {
    await runBrowserImeLifecycle(
      {
        update: async (text) => {
          await cdp.send("Input.imeSetComposition", {
            text,
            selectionStart: text.length,
            selectionEnd: text.length,
          });
          // An unsupported marked native proposal is refused while IME remains active.
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
            marked.textContent = paragraph.textContent;
            paragraph.replaceChildren(marked);
            editor.dispatchEvent(
              new InputEvent("input", {
                bubbles: true,
                inputType: "insertCompositionText",
                data: "契",
                isComposing: true,
              }),
            );
          });
          await page.waitForFunction(() =>
            globalThis.__folioCanonicalFuzzErrors?.some((error) =>
              error.message.includes("Composition changed unsupported content"),
            ),
          );
          // A full snapshot is unavailable while canonical composition remains active.
          expect(await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())).toBe(
            true,
          );
          expect(await page.locator(".ProseMirror").evaluate((editor) => editor.textContent)).toBe(
            "alpha",
          );
        },
        commit: async (text) => {
          await cdp.send("Input.insertText", { text });
          expect(cancelled).toBe(false);
          expect(await page.locator(".ProseMirror").getAttribute("data-final-prevented")).toBe(
            "true",
          );
          expect(await page.locator(".ProseMirror").getAttribute("data-native-end-delivered")).toBe(
            null,
          );
          // This assertion precedes cancel: removing product composition exit makes it fail.
          expect(await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())).toBe(
            false,
          );
          const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
          expect(current?.composing).toBe(false);
          expect(current?.text).toBe("alpha");
          expect(current?.canUndo).toBe(false);
          expect(current?.projectionMatchesCanonical).toBe(true);
        },
        cancel: async () => {
          cancelled = true;
          await cdp.send("Input.imeSetComposition", {
            text: "",
            selectionStart: 0,
            selectionEnd: 0,
          });
        },
      },
      { kind: "imeReplacement", updates: ["契"], completion: "commit" },
    );
    expect(cancelled).toBe(true);
  } finally {
    await cdp.detach();
  }
});

test("canonical history checks a lazy first view and rejects composition before reload", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  expect(await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())).toBe(null);
  const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
  const options = {
    page,
    source: [...new Uint8Array(source)],
    actions: [],
    missing: createMissingOpBurndown(),
  };
  expect(await checkCanonicalBrowserHistory(options)).toBe(0);
  expect(await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())).toBe(false);
  await page.locator(".ProseMirror").evaluate((element) => {
    element.dispatchEvent(new Event("compositionstart", { bubbles: true }));
  });
  await expect(checkCanonicalBrowserHistory(options)).rejects.toThrow(
    "case must start outside native composition",
  );
  // The failed invariant must not load/reset the owner and hide the leak.
  expect(await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())).toBe(true);
});

for (const { seed, path, kinds } of CANONICAL_BROWSER_HISTORY_REPLAYS) {
  test(`canonical history replay ${seed} ${path}`, async ({ page }) => {
    if (seed === 197 && path === "1") test.setTimeout(120_000);
    const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
    expect(traces).toHaveLength(1);
    const actions = traces.at(0);
    if (actions === undefined) throw new TypeError("Missing canonical regression trace");
    expect(actions.length).toBeGreaterThan(0);
    expect(actions.map(({ kind }) => kind)).toEqual(kinds);
    await page.goto("/?session=canonical");
    await page.waitForSelector(".layout-page");
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
    // Repeat identical package input to exercise adoption of a fresh owner.
    for (let load = 0; load < 2; load++) {
      const applied = await checkCanonicalBrowserHistory({
        page,
        source: [...new Uint8Array(source)],
        actions,
        missing: createMissingOpBurndown(),
      });
      expect(applied).toBeGreaterThan(0);
    }
  });
}
