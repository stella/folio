/**
 * A tracked replacement inside, or at the start of, another reviewer's record
 * (an insertion, or a run's pending formatting change) cuts that record and
 * puts the new text after the struck part. Whatever order its changes are
 * accepted and rejected in, the record keeps its id: accepting everything
 * gives the direct edit, rejecting everything gives the source back, and the
 * record's id never moves to a second record. The same holds when the
 * replacement inserts paragraphs, so the record's pieces end up apart.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertExactModel } from "../../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  type EditorIntent,
} from "../editorIntent";
import { identityKeysIn } from "../ids";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";

setDefaultTimeout(propertyTestTimeout(60_000));

const RECORD_ID = 13;
const TEXT = "abcdef";

const recordOf = (kind: "insertion" | "formatting"): ParagraphContent => {
  const run = (formatting: boolean): Run => ({
    type: "run",
    ...(formatting
      ? {
          formatting: { bold: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: RECORD_ID, author: "Earlier" },
              previousFormatting: {},
              currentFormatting: { bold: true },
            },
          ],
        }
      : {}),
    content: [{ type: "text", text: TEXT }],
  });
  return kind === "insertion"
    ? {
        type: "insertion",
        info: { id: RECORD_ID, author: "Earlier" },
        content: [run(false)],
      }
    : run(true);
};

const documentOf = (
  kind: "insertion" | "formatting",
  prefix: string,
): Document =>
  normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              ...(prefix === ""
                ? []
                : [
                    {
                      type: "run",
                      content: [{ type: "text", text: prefix }],
                    } satisfies Run,
                  ]),
              recordOf(kind),
            ],
          },
        ],
      },
    },
  });

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
): Document =>
  applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  }).unwrap().document;

/** How many records carry the earlier reviewer's id. */
const recordCount = (document: Document): number =>
  identityKeysIn(document.package.document.content).filter(
    (key) => key === identityKeysIn([recordOf("insertion")]).at(0),
  ).length;

test("a replacement keeps the id of another reviewer's record it cuts, in every resolution order", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("insertion" as const, "formatting" as const),
      fc.constantFrom("", "x"),
      fc.nat(TEXT.length - 1),
      fc.nat(TEXT.length - 1),
      fc.integer({ min: 0, max: 3 }),
      fc.array(fc.boolean(), { minLength: 8, maxLength: 8 }),
      fc.boolean(),
      (kind, prefix, start, span, paragraphs, decisions, reversed) => {
        const original = documentOf(kind, prefix);
        // From the record's start or inside it, to a gap inside it.
        const fromOffset = prefix.length + Math.min(start, TEXT.length - 2);
        const toOffset = Math.min(
          fromOffset +
            1 +
            (span % (TEXT.length - (fromOffset - prefix.length) - 1)),
          prefix.length + TEXT.length - 1,
        );
        const from = {
          story: OP_STORIES.MAIN,
          blockId: "00000001",
          offset: fromOffset,
        };
        const to = { ...from, offset: toOffset };
        const intent: EditorIntent =
          paragraphs === 0
            ? { type: "replaceText", from, to, text: "new" }
            : {
                type: "replaceFragment",
                from,
                to,
                paragraphs: Array.from({ length: paragraphs + 1 }, (_, i) => ({
                  type: "paragraph",
                  paraId: "00000001",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: `line${i}` }],
                    },
                  ],
                })),
                openStart: 1,
                openEnd: 1,
              };
        const allocation = allocateEditorIntentIds(original, intent);
        const direct = applyDocumentOps(
          original,
          compileEditorIntent(original, {
            intent,
            mode: { type: "editing", newIds: allocation.newIds },
          }).unwrap().ops,
        ).unwrap().document;
        const tracked = applyDocumentOps(
          original,
          compileEditorIntent(original, {
            intent,
            mode: {
              type: "suggesting",
              revision: {
                id: allocation.revisionId,
                author: "Reviewer",
                date: "2026-10-02T00:00:00Z",
              },
              newIds: allocation.newIds,
            },
          }).unwrap().ops,
        ).unwrap();
        const ids = reversed
          ? tracked.revisions.toReversed()
          : [...tracked.revisions];
        expect(recordCount(tracked.document)).toBe(1);
        // Every change on its own, each accepted or rejected as drawn.
        let mixed = tracked.document;
        for (const [index, id] of ids.entries()) {
          mixed = resolve(
            mixed,
            [id],
            decisions[index % decisions.length] === true
              ? REVISION_DECISIONS.ACCEPT
              : REVISION_DECISIONS.REJECT,
          );
          expect(recordCount(mixed)).toBe(1);
        }
        for (const decision of [
          REVISION_DECISIONS.ACCEPT,
          REVISION_DECISIONS.REJECT,
        ]) {
          let separately = tracked.document;
          for (const id of ids)
            separately = resolve(separately, [id], decision);
          const together = resolve(
            tracked.document,
            tracked.revisions,
            decision,
          );
          assertExactModel(separately, together);
          if (decision === REVISION_DECISIONS.REJECT)
            assertExactModel(together, original);
          else
            assertFreshIdentityEquivalent({
              actual: together,
              expected: direct,
              original,
              allocated: allocation.newIds,
            });
          expect(recordCount(together)).toBe(1);
        }
      },
    ),
    {
      numRuns: 200,
      id: "a replacement keeps the id of another reviewer's record it cuts, in every resolution order",
    },
  );
});
