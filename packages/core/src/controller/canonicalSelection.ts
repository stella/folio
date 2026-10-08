import { panic } from "better-result";
import { AllSelection, NodeSelection, TextSelection, type Transaction } from "prosemirror-state";
import type { CanonicalProjection, CanonicalSelection } from "./canonicalSession";

type RestoreCanonicalSelectionOptions = {
  transaction: Transaction;
  projection: CanonicalProjection;
  selection: CanonicalSelection;
  unavailable: { type: "refuse" } | { type: "near"; anchor: number };
};

/** Restore the admitted selection kind after a canonical projection changes. */
export const restoreCanonicalSelection = ({
  transaction,
  projection,
  selection,
  unavailable,
}: RestoreCanonicalSelectionOptions): boolean => {
  switch (selection.type) {
    case "inlineNode": {
      const anchor = projection.positionAt(selection.anchor);
      const head = projection.positionAt(selection.head);
      if (anchor.isErr() || head.isErr()) return false;
      const node = transaction.doc.nodeAt(anchor.value);
      if (
        !node?.isInline ||
        !node.isAtom ||
        !NodeSelection.isSelectable(node) ||
        anchor.value + node.nodeSize !== head.value
      )
        return false;
      transaction.setSelection(NodeSelection.create(transaction.doc, anchor.value));
      return true;
    }
    case "all":
      transaction.setSelection(new AllSelection(transaction.doc));
      return true;
    case "text": {
      const anchor = projection.positionAt(selection.anchor);
      const head = projection.positionAt(selection.head);
      if (anchor.isOk() && head.isOk()) {
        transaction.setSelection(TextSelection.create(transaction.doc, anchor.value, head.value));
        return true;
      }
      switch (unavailable.type) {
        case "refuse":
          return false;
        case "near":
          transaction.setSelection(
            TextSelection.near(
              transaction.doc.resolve(Math.min(unavailable.anchor, transaction.doc.content.size)),
            ),
          );
          return true;
        default: {
          const exhaustive: never = unavailable;
          return panic(`Unknown selection restoration policy: ${exhaustive}`);
        }
      }
    }
    default: {
      const exhaustive: never = selection.type;
      return panic(`Unknown canonical selection type: ${exhaustive}`);
    }
  }
};
