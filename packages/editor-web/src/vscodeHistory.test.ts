/**
 * The VS Code webview entry against the extension's real document session:
 * the workbench's undo stack and dirty state come from the steps the editor
 * reports, and its undo and redo come back as host messages. Every undo must
 * restore the document as it was before that step, every redo the document
 * after it, and the workbench's dirty state must follow.
 */

import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, setDefaultTimeout, test } from "bun:test";

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxSession } from "../../../plugins/vscode/src/session";
import type { HostMessage } from "./protocol";

setDefaultTimeout(60_000);

afterAll(() => GlobalRegistrator.unregister());

const fail = (message: string): never => {
  throw new Error(message);
};

/** Longer than the editor's undo grouping window. */
const NEW_UNDO_STEP_PAUSE_MS = 700;

const pause = (ms: number) =>
  new Promise((resolve) => {
    setTimeout(resolve, ms);
  });

const waitFor = async (condition: () => boolean, what: string) => {
  for (let poll = 0; poll < 400; poll += 1) {
    if (condition()) return;
    await pause(25);
  }
  fail(`Timed out waiting for ${what}.`);
};

let edits = 0;
const editCount = () => edits;
let loaded = 0;
const session = new DocxSession({
  fileName: "history.docx",
  documentPath: "/workspace/history.docx",
  initial: {
    bytes: new Uint8Array(
      await createDocx(createEmptyDocument({ initialText: "Start" })),
    ),
    baseline: "v1",
    restored: false,
  },
  author: "Test Author",
  locale: "en",
  mode: "editing",
  disk: {
    read: () => fail("The test never reads the disk."),
    versionOf: () => Promise.resolve(null),
    write: () => fail("The test never writes the disk."),
  },
  saveWithCli: () => fail("The test never saves."),
  describeChange: (version) => Promise.resolve(version),
  lease: null,
  ui: {
    confirmRewrite: () => Promise.resolve(undefined),
    resolveStale: () => Promise.resolve(undefined),
    resolveConflict: () => Promise.resolve(undefined),
    saveCopy: () => Promise.resolve(false),
    saveThroughWorkbench: () => Promise.resolve(false),
    revertThroughWorkbench: () => Promise.resolve(),
    notifyUpdated: () => {},
    showInfo: () => {},
    showLoadFailed: (message) => fail(`The document did not load: ${message}`),
    showError: (message) => fail(`The editor reported: ${message}`),
    showLeaseLost: () => {},
  },
  rewrite: {
    enabled: () => false,
    disable: () => Promise.resolve(),
    confirmed: false,
  },
  onEdit: () => {
    edits += 1;
  },
});

// The extension side: what the webview posts reaches the session, and what
// the session posts reaches the webview as a window message.
Reflect.set(globalThis, "acquireVsCodeApi", () => ({
  postMessage: (message: unknown) => {
    if (
      typeof message === "object" &&
      message !== null &&
      Reflect.get(message, "type") === "loaded"
    )
      loaded += 1;
    session.handleMessage(message);
  },
}));
session.attach((message: HostMessage) => window.postMessage(message, "*"));
await import("./vscode");

/** The body as the workbench would save it now. */
const savedBody = async () => {
  const { bytes } = await session.serialize();
  return (await parseDocx(bytes)).package.document;
};

const bodyView = (): HTMLElement =>
  document.querySelector<HTMLElement>(
    ".paged-editor__hidden-pm .ProseMirror",
  ) ?? fail("Expected the body editor.");

const type = (text: string) => {
  for (const character of text) {
    bodyView().dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: character,
      }),
    );
  }
};

const paste = (text: string) => {
  const clipboardData = new DataTransfer();
  clipboardData.setData("text/plain", text);
  bodyView().dispatchEvent(
    new ClipboardEvent("paste", {
      bubbles: true,
      cancelable: true,
      clipboardData,
    }),
  );
};

test("the workbench's undo and redo walk the editor's steps exactly, with dirty state following", async () => {
  await waitFor(() => loaded === 1, "the document to load");
  await waitFor(
    () =>
      document.querySelector(".paged-editor__hidden-pm .ProseMirror") !== null,
    "the body editor",
  );
  const states = [await savedBody()];
  const steps = [
    { name: "typing", run: () => type("Alpha") },
    {
      name: "a tracked edit",
      run: () => {
        session.setMode("suggesting");
        return pause(50).then(() => type("Beta"));
      },
    },
    { name: "a paste", run: () => paste("Gamma") },
  ] as const;
  for (const [index, step] of steps.entries()) {
    await pause(NEW_UNDO_STEP_PAUSE_MS);
    await step.run();
    const opened = index + 1;
    await waitFor(
      () => editCount() === opened,
      `${step.name} to open one step`,
    );
    expect(session.dirty).toBe(true);
    states.push(await savedBody());
    expect(states.at(-1)).not.toEqual(states.at(-2));
  }

  for (let index = steps.length - 1; index >= 0; index -= 1) {
    session.undo();
    await pause(100);
    expect(await savedBody()).toEqual(states[index] ?? fail("Missing state"));
    expect(session.dirty).toBe(index > 0);
  }
  for (let index = 1; index <= steps.length; index += 1) {
    session.redo();
    await pause(100);
    expect(await savedBody()).toEqual(states[index] ?? fail("Missing state"));
    expect(session.dirty).toBe(true);
  }
  // Undo and redo are the workbench's own; they open no further steps.
  expect(edits).toBe(steps.length);
});
