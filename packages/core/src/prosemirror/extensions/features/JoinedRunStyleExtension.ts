/**
 * Runs that a join moves into another paragraph take that paragraph's style.
 *
 * Deleting across a paragraph break (Backspace or Delete at the break, a
 * deletion or typing over a selection that spans paragraphs) runs the later
 * paragraph's remaining text on into the earlier one. A run's direct
 * formatting is its own and stays; what the later paragraph's style lent it
 * does not travel, and it reads in the style of the paragraph it now belongs
 * to. The document styles plugin re-reads those runs after an ordinary edit.
 *
 * A transaction that re-reads the runs it moves itself (resolving a tracked
 * change, an operation that merges paragraphs) says so with
 * {@link JOINED_RUNS_RESTYLED_META}, and is left as it is.
 */

import { Plugin, PluginKey } from "prosemirror-state";

import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";

export const JOINED_RUNS_RESTYLED_META = "joinedRunsRestyled";

const pastedTableEndKey = new PluginKey("pastedTableEnd");

/**
 * A table pasted over a story's last paragraph leaves that paragraph, emptied,
 * after it: a story ends with a paragraph, and the paste replaced the words,
 * not the break that ends the story. ProseMirror's paste replaces the whole
 * paragraph, so the empty one is put back.
 */
const createPastedTableEndPlugin = (): Plugin =>
  new Plugin({
    key: pastedTableEndKey,
    appendTransaction(transactions, oldState, newState) {
      const pasted = transactions.some(
        (transaction) => transaction.docChanged && transaction.getMeta("uiEvent") === "paste",
      );
      if (!pasted) {
        return null;
      }
      const last = newState.doc.lastChild;
      if (
        last?.type.spec["tableRole"] !== "table" ||
        oldState.doc.lastChild?.type.name !== "paragraph"
      ) {
        return null;
      }
      const paragraph = newState.schema.nodes["paragraph"]?.createAndFill();
      return paragraph ? newState.tr.insert(newState.doc.content.size, paragraph) : null;
    },
  });

export const JoinedRunStyleExtension = createExtension({
  name: "joinedRunStyle",
  defaultOptions: {},
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: [createPastedTableEndPlugin()],
    };
  },
});
