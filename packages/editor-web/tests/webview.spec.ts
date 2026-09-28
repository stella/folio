/**
 * The built VS Code bundle, end to end, in a page that carries the webview's
 * Content Security Policy and a fake extension on the other side of
 * `postMessage`. It checks what only a real browser shows: the policy admits
 * the bundle (no eval, no worker, no fetch), the fonts load, keys reach the
 * right owner, and the protocol round-trips a load, edits, undo, a save and a
 * reload.
 */

import { expect, test } from "@playwright/test";
import type { Page, Request } from "@playwright/test";
import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import path from "node:path";

import type { EditorMessage, HostMessage } from "../src/protocol";
import { ORIGIN, openWebview, postDocument, waitForSent } from "./webviewHarness";

const FIXTURE = path.resolve(import.meta.dirname, "../../playground/public/folio-showcase.docx");

const sent = (page: Page) => page.evaluate(() => globalThis.__folio.sent);

const sentOfType = async <Type extends EditorMessage["type"]>(page: Page, type: Type) =>
  (await sent(page)).filter(
    (message): message is Extract<EditorMessage, { type: Type }> => message.type === type,
  );

/** The bytes of the `serialized` answer to `requestId`. */
const serializedBytes = async (page: Page, requestId: number) =>
  new Uint8Array(
    await page.evaluate((id) => {
      const answer = globalThis.__folio.sent.find(
        (message) => message.type === "serialized" && message.requestId === id,
      );
      return answer?.type === "serialized" ? [...answer.bytes] : [];
    }, requestId),
  );

/** Post a command, as the extension's `webview.postMessage` does. */
const postToWebview = (page: Page, message: Exclude<HostMessage, { document: unknown }>) =>
  page.evaluate((command) => window.postMessage(command, "*"), message);

const paintedText = (page: Page) =>
  page.evaluate(() =>
    [...document.querySelectorAll(".layout-page")].map((node) => node.textContent).join("\n"),
  );

const MOD = process.platform === "darwin" ? "Meta" : "Control";
/** Longer than the editor's 500 ms history grouping window. */
const NEW_UNDO_STEP_PAUSE_MS = 700;

test("the VS Code bundle round-trips a document under the webview policy", async ({ page }) => {
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const fontRequests: Request[] = [];
  page.on("console", (message) => {
    if (message.type() === "error") consoleErrors.push(message.text());
  });
  page.on("pageerror", (error) => pageErrors.push(error.message));
  page.on("request", (request) => {
    if (request.resourceType() === "font") fontRequests.push(request);
  });

  await openWebview(page);
  await waitForSent(page, { type: "ready" });

  // A file that is not a .docx fails to load; the next load still succeeds.
  await postDocument(page, {
    type: "load",
    document: { bytes: new Uint8Array([1, 2, 3]), fileVersion: "v0", fileName: "broken.docx" },
    author: "Test Author",
    mode: "editing",
    locale: "en",
  });
  await waitForSent(page, { type: "loadFailed" });
  expect(await sentOfType(page, "loaded")).toEqual([]);

  const fixture = new Uint8Array(await readFile(FIXTURE));
  await postDocument(page, {
    type: "load",
    document: { bytes: fixture, fileVersion: "v1", fileName: "folio-showcase.docx" },
    author: "Test Author",
    mode: "editing",
    locale: "en",
  });
  await waitForSent(page, { type: "loaded", fileVersion: "v1" });
  await page.waitForFunction(() => document.querySelectorAll(".layout-page").length > 0);
  await page.evaluate(() => document.fonts.ready);

  // Type three undo steps: a word, a paragraph break, a second word.
  await page.locator(".layout-paragraph").first().click();
  await page.keyboard.press("End");
  await page.keyboard.type(" Alpha");
  await page.waitForTimeout(NEW_UNDO_STEP_PAUSE_MS);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(NEW_UNDO_STEP_PAUSE_MS);
  await page.keyboard.type("Beta");
  await page.waitForTimeout(NEW_UNDO_STEP_PAUSE_MS);
  await expect.poll(() => paintedText(page)).toContain("Beta");
  expect(await sentOfType(page, "edit")).toHaveLength(3);
  expect((await sentOfType(page, "dirty")).at(-1)).toEqual({ type: "dirty", dirty: true });

  // The undo keys belong to the host: pressed in the page, they change nothing.
  await page.keyboard.press(`${MOD}+z`);
  await page.waitForTimeout(100);
  expect(await paintedText(page)).toContain("Beta");
  // So does print: the editor opens no print window of its own.
  let popups = 0;
  page.on("popup", () => {
    popups += 1;
  });
  await page.keyboard.press(`${MOD}+p`);
  await page.waitForTimeout(200);
  expect(popups).toBe(0);

  // The host's undo walks back one step per message; redo walks forward.
  await postToWebview(page, { type: "undo" });
  await expect.poll(() => paintedText(page)).not.toContain("Beta");
  expect(await paintedText(page)).toContain("Alpha");
  await postToWebview(page, { type: "undo" });
  await postToWebview(page, { type: "undo" });
  await expect.poll(() => paintedText(page)).not.toContain("Alpha");
  for (let step = 0; step < 3; step += 1) await postToWebview(page, { type: "redo" });
  await expect.poll(() => paintedText(page)).toContain("Beta");
  expect(await sentOfType(page, "edit")).toHaveLength(3);

  // Save: a paragraph break forces a full repack.
  await postToWebview(page, { type: "serialize", requestId: 1 });
  await waitForSent(page, { type: "serialized", requestId: 1 });
  const [firstSave] = await sentOfType(page, "serialized");
  expect(firstSave?.fileVersion).toBe("v1");
  expect(firstSave?.strategy).toEqual({ type: "full-repack", reason: "structuralChange" });
  const savedBytes = await serializedBytes(page, 1);
  const savedZip = await JSZip.loadAsync(savedBytes);
  const savedXml = await savedZip.file("word/document.xml")?.async("text");
  expect(savedXml).toContain("Alpha");
  expect(savedXml).toContain("Beta");
  expect((await sentOfType(page, "dirty")).at(-1)).toEqual({ type: "dirty", dirty: false });

  // Suggesting mode from the host: typed text becomes a tracked insertion.
  await postToWebview(page, { type: "setMode", mode: "suggesting" });
  await page.waitForTimeout(NEW_UNDO_STEP_PAUSE_MS);
  await page.keyboard.type("Gamma");
  await expect.poll(() => page.locator(".layout-page .docx-insertion").count()).toBeGreaterThan(0);
  expect(await sentOfType(page, "edit")).toHaveLength(4);

  // A text-only edit keeps the selective patch in play.
  await postToWebview(page, { type: "serialize", requestId: 2 });
  await waitForSent(page, { type: "serialized", requestId: 2 });
  expect((await sentOfType(page, "serialized")).at(-1)?.strategy).toEqual({
    type: "selective-first",
  });

  // The toolbar's track-changes toggle reports the mode back to the host.
  await page.getByRole("button", { name: "Toggle track changes" }).click();
  await waitForSent(page, { type: "modeChanged", mode: "editing" });

  // Reload the saved bytes: the document is replaced and nothing counts as an edit.
  await postDocument(page, {
    type: "reload",
    document: { bytes: savedBytes, fileVersion: "v2", fileName: "folio-showcase.docx" },
  });
  await waitForSent(page, { type: "loaded", fileVersion: "v2" });
  await expect.poll(() => paintedText(page)).not.toContain("Gamma");
  expect(await paintedText(page)).toContain("Beta");
  expect(await sentOfType(page, "edit")).toHaveLength(4);

  // The dark theme switches folio's palette.
  await page.evaluate(() => document.body.classList.replace("vscode-light", "vscode-dark"));
  await expect(page.locator("html")).toHaveClass(/\bdark\b/u);

  // The policy admitted everything; fonts came from the bundle; no worker.
  const probe = await page.evaluate(() => globalThis.__folio);
  expect(probe.violations).toEqual([]);
  expect(probe.workers).toBe(0);
  expect(pageErrors).toEqual([]);
  expect(consoleErrors).toEqual([]);
  const loadedFaces = await page.evaluate(() =>
    [...document.fonts].filter((face) => face.status === "loaded").map((face) => face.family),
  );
  expect(loadedFaces.length).toBeGreaterThan(0);
  for (const request of fontRequests) {
    expect(request.url().startsWith(`${ORIGIN}/fonts/`)).toBe(true);
    expect((await request.response())?.status()).toBe(200);
  }
  const faces = [...new Set(loadedFaces)].join(", ");
  test.info().annotations.push({
    type: "fonts",
    description: `${String(fontRequests.length)} requests; loaded ${faces}`,
  });
});
