/**
 * Editor-side identity for one authored `w:r`, and the payload that element
 * carried which no formatting mark holds.
 *
 * Ordinary formatting marks cannot tell two adjacent source runs with
 * identical properties apart, so the save leg would fold them into one and
 * discard whichever record it did not keep. This mark is what keeps them
 * apart: it is in `getMarksKey`, so a change of identity is a run boundary,
 * and `joinsOwnedSourceRun` rejoins the leaves a page break cut out of one
 * authored run.
 *
 * `inclusive: false` settles the edges. The interior is settled by the strip
 * below: ProseMirror's stored marks at a position inside a span include every
 * mark of the preceding character, so text typed inside an authored run would
 * otherwise ride that run's `w:rsidR` — a claim about which editing session
 * wrote it. folio cannot register a session (`settings.xml`'s `w:rsids` is
 * copied through on every repack, and no save path mints one), so it must
 * never write a session id it did not read. Typed text becomes its own run
 * with no attributes instead, and `w:rsidRDefault` on the parent `w:p`
 * supplies the session the format itself says such a run belongs to.
 */

import { Plugin, PluginKey, type Transaction } from "prosemirror-state";

import { expectRunIdentityMarkAttrs } from "../../attrs";
import { nodesBetweenIndexed } from "../../indexedNodeLookup";
import { sweepPositions, type PositionQuery } from "../../positionSweep";
import {
  RUN_IDENTITY_ATTRIBUTE,
  RUN_IDENTITY_MARK_NAME,
  runIdentityAttrs,
} from "../../runIdentity";
import { createMarkExtension } from "../create";
import type { ExtensionRuntime } from "../types";

const CANONICAL_RUN_IDENTITY_ID = /^(?:0|[1-9]\d*)$/u;

const runIdentityStripKey = new PluginKey("runIdentityStrip");

type InsertedRange = { from: number; to: number };

/**
 * Every range `transactions` inserted, in the document they produced.
 *
 * A step map reports each replacement as an old span and the new span it
 * became; a new span longer than nothing is content the step put there, which
 * is exactly the text that may not claim an authored run. Each range is then
 * carried through the steps that followed it — the rest of its own
 * transaction, then every later one — so a batch applied together reports
 * positions in the document this hook is looking at rather than in an
 * intermediate one.
 *
 * @internal Exported for the per-step equivalence test.
 */
export const insertedRanges = (transactions: readonly Transaction[]): InsertedRange[] => {
  // The ranges are carried together, each entering after the step that made
  // it: slicing the mapping per step would walk the rest of the batch once
  // per step.
  const queries: PositionQuery[] = [];
  let stepOffset = 0;
  for (const transaction of transactions) {
    const offset = stepOffset;
    transaction.mapping.maps.forEach((stepMap, stepIndex) => {
      const from = offset + stepIndex + 1;
      stepMap.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        if (newEnd <= newStart) {
          return;
        }
        queries.push({ pos: newStart, assoc: -1, from }, { pos: newEnd, assoc: 1, from });
      });
    });
    stepOffset += transaction.mapping.maps.length;
  }
  const mapped = sweepPositions(
    transactions.map((transaction) => transaction.mapping),
    queries,
  );
  const ranges: InsertedRange[] = [];
  for (let index = 0; index < mapped.length; index += 2) {
    // SAFETY: every range queued its start and end together.
    ranges.push({ from: mapped[index]!.pos, to: mapped[index + 1]!.pos });
  }
  return ranges;
};

/**
 * Remove `runIdentity` from what a transaction inserted.
 *
 * Runs after the edit rather than in front of it because an insertion's marks
 * are decided by ProseMirror from the text around it, and there is no hook
 * that sees every route to that decision: typing, paste, a command and a
 * collaborative apply all arrive here and nowhere else in common.
 */
const createRunIdentityStripPlugin = (): Plugin =>
  new Plugin({
    key: runIdentityStripKey,
    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      const markType = newState.schema.marks[RUN_IDENTITY_MARK_NAME];
      if (markType === undefined) {
        return null;
      }
      const ranges = insertedRanges(transactions);
      if (ranges.length === 0) {
        return null;
      }
      const tr = newState.tr;
      const stripRange = ({ from, to }: InsertedRange): boolean => {
        const clampedTo = Math.min(to, newState.doc.content.size);
        const clampedFrom = Math.min(from, clampedTo);
        if (clampedFrom === clampedTo) {
          return false;
        }
        let rangeStripped = false;
        nodesBetweenIndexed(newState.doc, clampedFrom, clampedTo, (node, pos) => {
          if (!node.isInline || !markType.isInSet(node.marks)) {
            return true;
          }
          tr.removeMark(
            Math.max(pos, clampedFrom),
            Math.min(pos + node.nodeSize, clampedTo),
            markType,
          );
          rangeStripped = true;
          return true;
        });
        return rangeStripped;
      };
      let stripped = false;
      for (const range of ranges) {
        stripped = stripRange(range) || stripped;
      }
      if (!stripped) {
        return null;
      }
      // The strip is bookkeeping over what the user just wrote, not an edit of
      // its own: undo must return to the text before the insertion, not to the
      // same text still claiming an authored run.
      tr.setMeta("addToHistory", false);
      tr.setMeta(runIdentityStripKey, "stripped");
      if (newState.storedMarks !== null) {
        tr.setStoredMarks(newState.storedMarks.filter((mark) => mark.type !== markType));
      }
      return tr;
    },
  });

export const RunIdentityExtension = createMarkExtension({
  name: RUN_IDENTITY_MARK_NAME,
  schemaMarkName: RUN_IDENTITY_MARK_NAME,
  markSpec: {
    attrs: {
      id: {},
      preservedAttributes: { default: null },
      preserved: { default: null },
      emptyFormatting: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: `span[${RUN_IDENTITY_ATTRIBUTE}]`,
        getAttrs(dom) {
          const rawId = dom.dataset["docxRunIdentity"];
          if (rawId === undefined || !CANONICAL_RUN_IDENTITY_ID.test(rawId)) {
            return false;
          }
          const id = Number(rawId);
          return Number.isSafeInteger(id) ? runIdentityAttrs(id) : false;
        },
      },
    ],
    // The id and nothing else. An rsid names a session listed in the source
    // package's `settings.xml`, and folio cannot merge rsid tables, so a span
    // pasted into another document must not claim the session it came from. It
    // keys differently from its source for that reason and becomes its own run
    // with no attributes, which is the answer typing already gets.
    toDOM(mark) {
      const { id } = expectRunIdentityMarkAttrs(mark);
      return ["span", { [RUN_IDENTITY_ATTRIBUTE]: String(id) }, 0];
    },
  },
  onSchemaReady(): ExtensionRuntime {
    return { plugins: [createRunIdentityStripPlugin()] };
  },
});
