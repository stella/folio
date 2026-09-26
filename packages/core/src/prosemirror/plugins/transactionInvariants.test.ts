import { afterEach, describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { EditorState, Plugin } from "prosemirror-state";

import { documentShape } from "../../__tests__/documentShapes";
import { createHarnessState, parseShapeDocument } from "../../__tests__/editorHarness";
import { ExtensionManager } from "../extensions/ExtensionManager";
import { createStarterKit } from "../extensions/StarterKit";
import { toggleNumberedList } from "../extensions/features/ListExtension";
import { singletonManager } from "../schema";
import {
  checkEditorStateInvariants,
  createTransactionInvariantPlugin,
  TransactionInvariantError,
} from "./transactionInvariants";

const withInvariantPlugin = async (shapeId: string): Promise<EditorState> => {
  const document = await parseShapeDocument(await documentShape(shapeId).build());
  return createHarnessState(document, "editing", [createTransactionInvariantPlugin()]);
};

const caretIn = (state: EditorState, text: string): EditorState => {
  let target = -1;
  state.doc.descendants((node, pos) => {
    if (target === -1 && node.isTextblock && node.textContent.includes(text)) {
      target = pos + 1;
    }
    return target === -1;
  });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, target)));
};

const hasInvariantPlugin = (plugins: readonly Plugin[]): boolean =>
  plugins.some((plugin) =>
    (plugin as unknown as { key: string }).key.startsWith("transactionInvariants"),
  );

describe("transaction invariants", () => {
  const flags = globalThis.__folioTransactionInvariants;
  afterEach(() => {
    globalThis.__folioTransactionInvariants = flags;
  });

  test("a transaction that references an undefined numbering instance throws with its steps", async () => {
    // A package with no numbering part: the list command names an instance
    // nothing defines (#1091), which the save path would refuse.
    const state = caretIn(await withInvariantPlugin("plain-markdown"), "First item");

    let thrown: unknown;
    try {
      toggleNumberedList(state, (tr) => state.apply(tr));
    } catch (error) {
      thrown = error;
    }

    expect(thrown).toBeInstanceOf(TransactionInvariantError);
    const error = thrown as TransactionInvariantError;
    expect(error.message).toContain("Numbering definition 2 is missing.");
    expect(error.steps.length).toBeGreaterThan(0);
    expect(error.message).toContain('"stepType"');
  });

  test("a sound transaction passes", async () => {
    const state = caretIn(await withInvariantPlugin("single-decimal-list"), "Plain text.");

    let next: EditorState | null = null;
    toggleNumberedList(state, (tr) => {
      next = state.apply(tr);
    });

    expect(next).not.toBeNull();
    expect(checkEditorStateInvariants(next ?? state)).toEqual([]);
  });

  test("the extension manager installs the plugin only while the flag is set", () => {
    globalThis.__folioTransactionInvariants = undefined;
    const off = new ExtensionManager(createStarterKit());
    off.buildSchema();
    off.initializeRuntime();
    expect(hasInvariantPlugin(off.getPlugins())).toBe(false);

    globalThis.__folioTransactionInvariants = { enabled: true };
    const on = new ExtensionManager(createStarterKit());
    on.buildSchema();
    on.initializeRuntime();
    expect(hasInvariantPlugin(on.getPlugins())).toBe(true);
  });

  test("the test preload enables it for editor runtimes", () => {
    expect(hasInvariantPlugin(singletonManager.getPlugins())).toBe(true);
  });
});
