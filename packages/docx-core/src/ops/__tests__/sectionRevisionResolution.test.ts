import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const info = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-11-20T00:00:00Z",
});

const text = (paraId: string, value: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: value }] }],
});

const headerReferences = [{ type: "default" as const, rId: "rIdNow" }];

/** A body whose first section's end carries a tracked section change. */
const sectionDocument = (): Document => ({
  package: {
    document: {
      content: [
        text("00000001", "Opening"),
        {
          ...text("00000002", "Section end"),
          sectionProperties: {
            sectionStart: "nextPage",
            headerReferences,
            propertyChanges: [
              {
                type: "sectionPropertyChange",
                info: info(7),
                previousProperties: { sectionStart: "continuous" },
              },
            ],
          },
        },
        {
          type: "paragraph",
          paraId: "00000003",
          content: [
            {
              type: "insertion",
              info: info(8),
              content: [
                { type: "run", content: [{ type: "text", text: "Added" }] },
              ],
            },
          ],
        },
      ],
    },
  },
});

const sectionEnd = (document: Document) => {
  const block = document.package.document.content.at(1);
  if (block?.type !== "paragraph")
    throw new Error("The section end stays a paragraph.");
  return block.sectionProperties;
};

const resolve = (revisionIds: number[], decision: RevisionDecision) => {
  const before = sectionDocument();
  const result = applyDocumentOp(before, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (result.isErr()) throw result.error;
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toEqual(before);
  return result.value.document;
};

test("accepting a body section's change keeps its properties", () => {
  expect(sectionEnd(resolve([7], REVISION_DECISIONS.ACCEPT))).toEqual({
    sectionStart: "nextPage",
    headerReferences,
  });
});

test("rejecting a body section's change restores its properties and keeps its references", () => {
  expect(sectionEnd(resolve([7], REVISION_DECISIONS.REJECT))).toEqual({
    sectionStart: "continuous",
    headerReferences,
  });
});

for (const decision of Object.values(REVISION_DECISIONS))
  test(`a section change resolves with the inline changes beside it (${decision})`, () => {
    const resolved = resolve([7, 8], decision);
    expect(sectionEnd(resolved)?.propertyChanges).toBeUndefined();
    expect(JSON.stringify(resolved).includes("Added")).toBe(
      decision === REVISION_DECISIONS.ACCEPT,
    );
    expect(JSON.stringify(resolved)).not.toContain('"insertion"');
  });
