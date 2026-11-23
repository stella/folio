import { panic } from "better-result";
import { Plugin, PluginKey } from "prosemirror-state";

import type { FolioDerivedFields } from "../../types/fields";

const REMOVAL_MARKS = new Set(["deletion", "moveFrom"]);

/**
 * The default session's derived fields: after each transaction that changes
 * the document, the owned fields outside tracked deletions show the texts the
 * host derives for them, in the same transaction as the edit.
 */
export const createDerivedFieldsPlugin = (
  getDerivedFields: () => FolioDerivedFields | undefined,
): Plugin =>
  new Plugin({
    // One key per editor, so two editors keep separate plugins.
    key: new PluginKey("folioDerivedFields"),
    appendTransaction(transactions, _previous, state) {
      const derived = getDerivedFields();
      if (
        derived === undefined ||
        !transactions.some((transaction) => transaction.docChanged)
      )
        return null;
      const owned: { position: number; displayText: string }[] = [];
      const instructions: string[] = [];
      state.doc.descendants((node, position) => {
        if (node.type.name !== "field") return true;
        const { instruction, displayText } = node.attrs;
        if (
          typeof instruction !== "string" ||
          typeof displayText !== "string" ||
          !derived.owns(instruction) ||
          node.marks.some((mark) => REMOVAL_MARKS.has(mark.type.name))
        )
          return false;
        owned.push({ position, displayText });
        instructions.push(instruction);
        return false;
      });
      if (owned.length === 0) return null;
      const texts = derived.displayTexts(instructions);
      if (texts.length !== owned.length)
        panic(
          `derivedFields.displayTexts returned ${texts.length} texts for ${owned.length} fields.`,
        );
      const transaction = state.tr;
      for (const [index, { position, displayText }] of owned.entries()) {
        const text = texts[index];
        if (text !== undefined && text !== displayText)
          transaction.setNodeAttribute(position, "displayText", text);
      }
      return transaction.docChanged ? transaction : null;
    },
  });
