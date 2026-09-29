/**
 * Runs that a join moves into another paragraph take that paragraph's style.
 *
 * Deleting across a paragraph break (Backspace or Delete at the break, a
 * deletion or typing over a selection that spans paragraphs) runs the later
 * paragraph's remaining text on into the earlier one. A run's direct
 * formatting is its own and stays; what the later paragraph's style lent it
 * does not travel, and it reads in the style of the paragraph it now belongs
 * to — the editor's marks hold inherited formatting too, so without this the
 * moved runs would save the old style's look as direct formatting.
 *
 * A transaction that re-reads the runs it moves itself (resolving a tracked
 * change, an operation that merges paragraphs) says so with
 * {@link JOINED_RUNS_RESTYLED_META}.
 */

import { isHistoryTransaction } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey, type Transaction } from "prosemirror-state";
import { Mapping, ReplaceStep } from "prosemirror-transform";

import { getDocumentStyleResolver } from "../../plugins/documentStyles";
import { rebaseParagraphRuns } from "../../rebaseParagraphRunFormatting";
import { paragraphRunStyleContext } from "../../runStyleFormatting";
import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";

export const JOINED_RUNS_RESTYLED_META = "joinedRunsRestyled";

const joinedRunStyleKey = new PluginKey("joinedRunStyle");

type MovedRuns = {
  /** The paragraph the runs came from, before the join. */
  paragraph: PMNode;
  /** Where they start and end in the document after every transaction. */
  from: number;
  to: number;
};

/** Content of a later paragraph that a step ran on into an earlier one. */
const movedRunsOf = (transactions: readonly Transaction[]): MovedRuns[] => {
  const moved: MovedRuns[] = [];
  transactions.forEach((transaction, transactionIndex) => {
    // Undo and redo replay their event whole, this transaction's own restyle included.
    if (
      transaction.getMeta(JOINED_RUNS_RESTYLED_META) === true ||
      isHistoryTransaction(transaction)
    ) {
      return;
    }
    transaction.steps.forEach((step, stepIndex) => {
      if (!(step instanceof ReplaceStep)) {
        return;
      }
      const { from, to } = step as unknown as { from: number; to: number };
      const before = transaction.docs[stepIndex];
      if (!before || to <= from) {
        return;
      }
      const $from = before.resolve(from);
      const $to = before.resolve(to);
      if (
        $from.parent === $to.parent ||
        $from.parent.type.name !== "paragraph" ||
        $to.parent.type.name !== "paragraph" ||
        $to.end() <= to
      ) {
        return;
      }
      const mapping = new Mapping(transaction.mapping.maps.slice(stepIndex));
      for (const later of transactions.slice(transactionIndex + 1)) {
        mapping.appendMapping(later.mapping);
      }
      moved.push({
        paragraph: $to.parent,
        from: mapping.map(to, 1),
        to: mapping.map($to.end(), -1),
      });
    });
  });
  return moved;
};

const createJoinedRunStylePlugin = (): Plugin =>
  new Plugin({
    key: joinedRunStyleKey,
    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      const moved = movedRunsOf(transactions);
      if (moved.length === 0) {
        return null;
      }
      const styleResolver = getDocumentStyleResolver(newState);
      if (!styleResolver) {
        return null;
      }
      const tr = newState.tr;
      for (const { paragraph, from, to } of moved) {
        if (to <= from || to > tr.doc.content.size) {
          continue;
        }
        const $from = tr.doc.resolve(from);
        // Still the start of its own paragraph: nothing joined it.
        if ($from.parent.type.name !== "paragraph" || $from.parentOffset === 0) {
          continue;
        }
        if ($from.parent.sameMarkup(paragraph)) {
          continue;
        }
        rebaseParagraphRuns({
          previousContext: paragraphRunStyleContext(paragraph, styleResolver),
          paragraphPosition: $from.before(),
          range: { from: $from.parentOffset, to: Math.min(to, $from.end()) - $from.start() },
          styleResolver,
          tr,
        });
      }
      if (!tr.docChanged) {
        return null;
      }
      tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
      return tr;
    },
  });

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
      plugins: [createJoinedRunStylePlugin(), createPastedTableEndPlugin()],
    };
  },
});
