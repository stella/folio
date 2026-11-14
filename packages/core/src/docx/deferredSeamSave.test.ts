import { expect, setDefaultTimeout, test } from "bun:test";

import { validateDocxPackage } from "@stll/docx-core";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import { assertExactModel } from "../../../../test/exactModel";
import { propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, ParagraphContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

// A seam waiting for its pieces is recorded on the pending changes standing
// at it: which changes must go first, how deep it joins, and the date of the
// revision whose kept boundaries wait there. That record survives a save, so
// resolving the reopened file gives what resolving the open one does, in any
// order.

const link = (text: string): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentWith = (content: Paragraph["content"]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "00000001", content },
  ];
  return document;
};

/** A link piece in a pending deletion, a kept bookmark end, an insertion waiting with the seam, the other piece. */
const tracked = (): Document =>
  documentWith([
    {
      type: "deletion",
      info: { id: 20, author: "B", date: "2026-03-04T05:00:01Z" },
      content: [link("x")],
      resolutionJoins: { before: 0, after: 2, remove: 0 },
    },
    { type: "bookmarkEnd", id: 1 },
    {
      type: "insertion",
      info: { id: 6, author: "A", date: "2026-03-04T05:00:03Z" },
      content: [{ type: "run", content: [{ type: "text", text: "y" }] }],
      resolutionJoins: {
        before: 0,
        after: 0,
        remove: 0,
        deferredRemove: [
          { depth: 2, blockers: [6], keptFrom: "2026-03-04T05:00:02Z" },
        ],
      },
    },
    link("a"),
  ]);

const paragraph = (document: Document) => {
  const first = document.package.document.content[0];
  return first?.type === "paragraph" ? first.content : [];
};

/** The paragraph's records, text and boundaries, without package-level ids. */
const shape = (nodes: readonly unknown[]): unknown[] =>
  nodes.map((node) => {
    if (typeof node !== "object" || node === null) return node;
    const record: Record<string, unknown> = Object.fromEntries(
      Object.entries(node),
    );
    const children = record["content"] ?? record["children"];
    return {
      type: record["type"],
      ...(typeof record["text"] === "string" ? { text: record["text"] } : {}),
      ...(typeof record["id"] === "number" ? { id: record["id"] } : {}),
      ...(Array.isArray(children) ? { children: shape(children) } : {}),
    };
  });

const resolved = (
  document: Document,
  steps: readonly (readonly [number, RevisionDecision])[],
): Document => {
  let current = document;
  for (const [id, decision] of steps) {
    const applied = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [id],
      decision,
    });
    if (applied.isErr()) throw applied.error;
    current = applied.value.document;
  }
  return current;
};

test("a deferred seam with its kept date survives a save and resolves as before, in every order", async () => {
  const source = tracked();
  const saved = await repackDocx(source, { updateModifiedDate: false });
  expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({
    valid: true,
  });
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(paragraph(reopened)).toMatchObject([
    { type: "deletion", resolutionJoins: { before: 0, after: 2, remove: 0 } },
    { type: "bookmarkEnd", id: 1 },
    {
      type: "insertion",
      resolutionJoins: {
        deferredRemove: [
          { depth: 2, blockers: [6], keptFrom: "2026-03-04T05:00:02Z" },
        ],
      },
    },
    { type: "hyperlink" },
  ]);
  const { ACCEPT, REJECT } = REVISION_DECISIONS;
  for (const steps of [
    [
      [6, REJECT],
      [20, ACCEPT],
    ],
    [
      [20, ACCEPT],
      [6, REJECT],
    ],
    [
      [6, REJECT],
      [20, REJECT],
    ],
    [
      [20, REJECT],
      [6, REJECT],
    ],
    [
      [6, ACCEPT],
      [20, ACCEPT],
    ],
    [
      [20, REJECT],
      [6, ACCEPT],
    ],
  ] as const)
    assertExactModel(
      shape(paragraph(resolved(reopened, steps))),
      shape(paragraph(resolved(source, steps))),
    );
});
