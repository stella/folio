/**
 * A tracked replacement inside one source run leaves its payload between the
 * run's two cut pieces. Resolving the deletion and the insertion one at a
 * time, in either order and with either decision each, gives the document
 * resolving them together gives: the pieces rejoin once nothing new stands
 * between them, at top level and inside a content control.
 */
import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const documentOf = (inControl: boolean): Document => {
  const runs: ParagraphContent[] = [run("aaa"), run("bbb")];
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: inControl
              ? [
                  {
                    type: "inlineSdt",
                    properties: { id: 100, sdtType: "richText" },
                    content: runs,
                  },
                ]
              : runs,
          },
        ],
      },
    },
  };
};

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
});

const tracked = (document: Document, from: number, to: number) => {
  const compiled = compileEditorIntent(document, {
    intent: { type: "replaceText", from: at(from), to: at(to), text: "new" },
    mode: {
      type: "suggesting",
      revision: { id: 1, author: "E", date: "2026-01-01T00:00:00Z" },
      newIds: { revision: [20, 21, 22, 23], control: [200, 201] },
    },
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const revisionIds = (document: Document): number[] => [
  ...new Set(
    [...JSON.stringify(document).matchAll(/"info":\{"id":(\d+)/gu)].map(
      ([, id]) => Number(id),
    ),
  ),
];

const resolve = (
  document: Document,
  ids: readonly number[],
  decision: RevisionDecision,
): Document => {
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: ids,
      decision,
    },
  ]);
  if (resolved.isErr()) throw resolved.error;
  return resolved.value.document;
};

const DECISIONS = [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT];
// The middle of the first run, and its end, where the cut meets the next run.
const RANGES = [
  [1, 2],
  [1, 3],
] as const;

describe("a tracked replacement inside a source run", () => {
  for (const inControl of [false, true])
    for (const [from, to] of RANGES)
      test(`${inControl ? "inside a control" : "at top level"}, replacing ${from}..${to}: resolving one at a time equals resolving together`, () => {
        const pending = tracked(documentOf(inControl), from, to);
        const ids = revisionIds(pending);
        expect(ids).toHaveLength(2);
        for (const first of DECISIONS)
          for (const second of DECISIONS)
            for (const order of [ids, ids.toReversed()]) {
              const [one = 0, two = 0] = order;
              const decided = (id: number) => (id === one ? first : second);
              const inTurn = resolve(
                resolve(pending, [one], decided(one)),
                [two],
                decided(two),
              );
              const together =
                first === second
                  ? resolve(pending, ids, first)
                  : resolve(
                      resolve(pending, [two], decided(two)),
                      [one],
                      decided(one),
                    );
              expect(inTurn).toEqual(together);
            }
      });
});
