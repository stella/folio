import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { createEditorIntentIdAllocator } from "../editorIntent";
import { formatParaId, takeParagraphIds } from "../idLease";
import * as identities from "../ids";
import { OP_STORIES } from "../types";

test("allocation queries package membership and respects the lease and reserved spellings", () => {
  const document: Document = {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "0000000a", content: [] }],
      },
    },
  };
  const census = spyOn(identities, "packageParagraphIds");
  try {
    const options = {
      document,
      lease: { blocks: [{ kind: "paragraph", start: 10, end: 14 }] },
      reserved: new Set(["0000000c"]),
    } as const;
    expect(takeParagraphIds({ ...options, count: 2 })).toEqual([
      "0000000B",
      "0000000D",
    ]);
    expect(takeParagraphIds({ ...options, count: 3 })).toBeUndefined();
    expect(census).toHaveBeenCalledTimes(0);
  } finally {
    census.mockRestore();
  }
});

test("a dense occupied prefix follows immutable edits and reopens a deleted id", () => {
  const paragraph = (id: number): Paragraph => ({
    type: "paragraph",
    paraId: formatParaId(id).toLowerCase(),
    content: [],
  });
  const document = (content: Paragraph[]): Document => ({
    package: { document: { content } },
  });
  const content = Array.from({ length: 256 }, (_, index) =>
    paragraph(index + 1),
  );
  const source = document(content.slice());
  const reserved = new Set<string>();
  const options = { document: source, reserved, lease: undefined, count: 1 };
  expect(takeParagraphIds(options)).toEqual(["00000101"]);

  const probes = spyOn(identities.paragraphIdSequence, "keyOf");
  try {
    // Cross the key index's chain-settlement boundary.
    for (let id = 257; id <= 296; id += 1) {
      content.push(paragraph(id));
      expect(
        takeParagraphIds({ ...options, document: document(content.slice()) }),
      ).toEqual([formatParaId(id + 1)]);
    }
    // Each successor probes the added id and its next candidate.
    expect(probes).toHaveBeenCalledTimes(80);
  } finally {
    probes.mockRestore();
  }

  const afterDeletion = document(content.toSpliced(9, 1));
  expect(takeParagraphIds({ ...options, document: afterDeletion })).toEqual([
    "0000000A",
  ]);
  reserved.add("0000000a");
  expect(takeParagraphIds({ ...options, document: afterDeletion })).toEqual([
    "00000129",
  ]);
  expect(takeParagraphIds(options)).toEqual(["00000101"]);
});

test("revision allocation honors real records, pending reservations, and lease bounds", () => {
  const document = (physicalId: number, blockers: number[]): Document => ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "insertion",
                info: { id: physicalId, author: "Editor" },
                content: [
                  { type: "run", content: [{ type: "text", text: "x" }] },
                ],
                resolutionJoins: {
                  before: 0,
                  after: 0,
                  remove: 0,
                  deferredRemove: [{ depth: 0, blockers }],
                },
              },
            ],
          },
        ],
      },
    },
  });
  const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 };
  const intent = { type: "replaceText", from: at, to: at, text: "y" } as const;
  const allocate = createEditorIntentIdAllocator();
  const source = document(1, [2]);
  expect(allocate(source, intent).revisionId).toBe(3);

  const retired = document(3, [1, 2]);
  expect(allocate(retired, intent).revisionId).toBe(4);
  expect(
    allocate(retired, intent, {
      blocks: [{ kind: "revision", start: 5, end: 7 }],
    }),
  ).toMatchObject({ revisionId: 5, newIds: { revision: [6] } });

  expect(allocate(document(3, [1]), intent).revisionId).toBe(2);
  expect(allocate(source, intent).revisionId).toBe(3);
});
