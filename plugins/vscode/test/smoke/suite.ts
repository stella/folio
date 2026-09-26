// Runs inside VS Code (see run.ts): the extension as users get it, with its
// test hooks on (`FOLIO_VSCODE_TEST=1`). Opens the document, checks it is in
// folio's editor and the webview loaded it, types, saves, and checks the file
// on disk; then types again and reverts.

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import * as vscode from "vscode";

import type { EditorTestHooks } from "../../src/editor";

const EXTENSION_ID = "stll.stella-folio";
const EDITOR_VIEW_TYPE = "folio.docxEditor";
const TYPED = "Typed in VS Code";
const DISCARDED = "Reverted away";

const log = (line: string) => console.log(`[smoke] ${line}`);

const describeInput = (input: unknown): string => {
  if (input instanceof vscode.TabInputCustom) return `custom ${input.viewType} ${input.uri.fsPath}`;
  if (input instanceof vscode.TabInputText) return `text ${input.uri.fsPath}`;
  return "other";
};

/** Every open tab, for a failure message. */
const describeTabs = (): string =>
  vscode.window.tabGroups.all
    .flatMap((group) =>
      group.tabs.map(
        (tab) => `"${tab.label}" (${describeInput(tab.input)}${tab.isActive ? ", active" : ""})`,
      ),
    )
    .join("; ");

/** What the webview sent so far, for a failure message. */
let describeReceived = (): string => "nothing yet";

const waitFor = async (what: string, probe: () => boolean, timeoutMs = 60_000): Promise<void> => {
  const deadline = Date.now() + timeoutMs;
  while (!probe()) {
    if (Date.now() > deadline) {
      throw new Error(
        `Timed out waiting for ${what}. Open tabs: ${describeTabs() || "none"}. The webview sent: ${describeReceived()}.`,
      );
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  log(what);
};

const sha = (file: string) => createHash("sha256").update(readFileSync(file)).digest("hex");

/** The tab showing `uri` in folio's editor, in any group. */
const editorTab = (uri: vscode.Uri): vscode.Tab | undefined =>
  vscode.window.tabGroups.all
    .flatMap((group) => group.tabs)
    .find(
      ({ input }) =>
        input instanceof vscode.TabInputCustom &&
        input.viewType === EDITOR_VIEW_TYPE &&
        input.uri.toString() === uri.toString(),
    );

const field = (value: unknown, key: string): unknown =>
  typeof value === "object" && value !== null ? Reflect.get(value, key) : undefined;

const isType = (message: unknown, type: string) => field(message, "type") === type;

/** The document's text as the bundled CLI reads it from disk. */
const textOnDisk = (cliEntry: string, file: string): string => {
  const result = spawnSync(process.execPath, [cliEntry, "read", file, "--output", "json"], {
    encoding: "utf8",
    env: { ...process.env, ELECTRON_RUN_AS_NODE: "1" },
  });
  const envelope: unknown = JSON.parse(result.stdout.trim().split("\n").at(-1) ?? "null");
  const blocks = field(field(field(envelope, "data"), "result"), "blocks");
  assert.ok(Array.isArray(blocks), `folio read failed: ${result.stdout}${result.stderr}`);
  return blocks.map((block: unknown) => String(field(block, "text"))).join("\n");
};

export const run = async (): Promise<void> => {
  const documentPath = process.env["FOLIO_SMOKE_DOCUMENT"];
  assert.ok(documentPath !== undefined, "FOLIO_SMOKE_DOCUMENT");
  const uri = vscode.Uri.file(documentPath);
  const directory = path.dirname(documentPath);
  const name = path.basename(documentPath);
  const lockPath = path.join(directory, `.${name}.folio-lock`);
  const extension = vscode.extensions.getExtension<EditorTestHooks | undefined>(EXTENSION_ID);
  assert.ok(extension !== undefined, `${EXTENSION_ID} is loaded`);
  const cliEntry = path.join(extension.extensionPath, "dist", "cli", "folio.mjs");
  // A modal would hang the test; this test's edits are text-only anyway.
  await vscode.workspace
    .getConfiguration()
    .update("folio.editor.confirmRewrite", false, vscode.ConfigurationTarget.Global);

  // Opening a .docx lands in folio's editor, and the webview loads it.
  await vscode.commands.executeCommand("vscode.open", uri);
  await waitFor("the document opens in folio's editor", () => editorTab(uri) !== undefined);
  const activeTab = () => editorTab(uri);
  const hooks = extension.exports;
  assert.ok(hooks !== undefined, "the extension exports its test hooks");
  describeReceived = () =>
    JSON.stringify(
      hooks.received(uri).map((message) => {
        const type = field(message, "type");
        return type === "loadFailed" || type === "error" ? message : type;
      }),
    );
  const count = (type: string) =>
    hooks.received(uri).filter((message) => isType(message, type)).length;
  await waitFor(
    "the webview posted loaded",
    () => count("loaded") > 0 || count("loadFailed") > 0,
    180_000,
  );
  assert.equal(count("loadFailed"), 0, "the document loads");

  // Type: one edit, a dirty tab, and the editor lease on disk.
  const original = sha(documentPath);
  assert.ok(await hooks.post(uri, { type: "folio-test-type", text: ` ${TYPED}.` }));
  await waitFor("an edit", () => count("edit") > 0);
  await waitFor("the tab is dirty", () => activeTab()?.isDirty === true);
  await waitFor("the editor holds the lease", () => existsSync(lockPath));
  assert.match(readFileSync(lockPath, "utf8"), /"owner":"folio-vscode"/u);

  // Save: the file changed, holds the typed text, is backed up and journaled.
  const saved = await vscode.workspace.save(uri);
  assert.equal(saved?.toString(), uri.toString(), "the save went through");
  await waitFor("the tab is clean", () => activeTab()?.isDirty === false);
  const savedVersion = sha(documentPath);
  assert.notEqual(savedVersion, original, "the file on disk changed");
  assert.ok(textOnDisk(cliEntry, documentPath).includes(TYPED), "the typed text is on disk");
  assert.ok(readdirSync(path.join(directory, ".folio", "backups", name)).length > 0);
  const lastLine = readFileSync(path.join(directory, ".folio", "journal.jsonl"), "utf8")
    .trim()
    .split("\n")
    .at(-1);
  assert.match(lastLine ?? "", /"tool":"editor_save"/u);
  assert.match(lastLine ?? "", /"surface":"vscode"/u);
  await waitFor("the lease is released", () => !existsSync(lockPath));

  // Revert: type again, revert, and the webview reloads the saved file.
  const loads = count("loaded");
  const edits = count("edit");
  assert.ok(await hooks.post(uri, { type: "folio-test-type", text: ` ${DISCARDED}.` }));
  await waitFor("a second edit", () => count("edit") > edits);
  await waitFor("the tab is dirty again", () => activeTab()?.isDirty === true);
  await vscode.commands.executeCommand("workbench.action.files.revert");
  await waitFor("the tab is clean after revert", () => activeTab()?.isDirty === false);
  await waitFor("the webview reloaded", () => count("loaded") > loads);
  assert.equal(sha(documentPath), savedVersion, "revert leaves the file alone");
  assert.ok(!textOnDisk(cliEntry, documentPath).includes(DISCARDED));
  await waitFor("the lease is released after revert", () => !existsSync(lockPath));

  await vscode.commands.executeCommand("workbench.action.closeAllEditors");
  log("passed");
};
