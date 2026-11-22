import { expect, test } from "bun:test";

import { propertyTestTimeout } from "../../../../test/property-testing";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

// Every three-paragraph document that starts with words, each paragraph with
// or without words, with or without its mark pending deletion, numbered or
// not, and each of its paragraphs deleted: a tracked or suggested deletion,
// accepted with every other change, reads as the direct deletion does.

type Shape = { words: boolean; pending: boolean; numbered: boolean };

const BOOLEANS = [false, true] as const;
const SHAPES: readonly Shape[] = BOOLEANS.flatMap((words) =>
  BOOLEANS.flatMap((pending) =>
    BOOLEANS.map((numbered) => ({ words, pending, numbered })),
  ),
);
const DOCUMENTS = SHAPES.filter(({ words }) => words).flatMap((first) =>
  SHAPES.flatMap((second) =>
    SHAPES.filter(({ pending }) => !pending).map((last) => [
      first,
      second,
      last,
    ]),
  ),
);

const idOf = (index: number): string => `0000000${index + 1}`;

const documentOf = (shapes: readonly Shape[]): Document => ({
  package: {
    document: {
      content: shapes.map(({ words, pending, numbered }, index): Paragraph => ({
        type: "paragraph",
        paraId: idOf(index),
        content: words
          ? [
              {
                type: "run",
                content: [{ type: "text", text: `w${index + 1}` }],
              },
            ]
          : [],
        ...(numbered
          ? {
              formatting: {
                numPr: { kind: "reference", numId: 1, ilvl: 0 },
              },
            }
          : {}),
        ...(pending
          ? {
              pPrMark: {
                kind: "del",
                info: {
                  id: index + 1,
                  author: "Earlier",
                  date: "2026-01-02T03:04:05Z",
                },
              },
            }
          : {}),
      })),
    },
    numbering: {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    },
  },
});

/** The document's words and list markers once everything is accepted. */
const accepted = async (
  bytes: ArrayBuffer,
  blockId: string,
  mode: "direct" | "tracked-changes" | "suggested",
): Promise<string | undefined> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    author: "Reviewer",
    session: "canonical",
  });
  const { applied } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "op", type: "deleteBlock", blockId }],
    },
    { undefinedReferences: "refuse" },
  );
  if (applied.length === 0) return undefined;
  for (const id of new Set(
    reviewer.exportPendingSuggestions().map(({ suggestionId }) => suggestionId),
  ))
    expect(reviewer.acceptSuggestion(id)).toBe(true);
  reviewer.acceptAll();
  return reviewer
    .getContentAsText()
    .replace(/^\[[^\]\n]*\]/gmu, "")
    .replace(/\s+/gu, "");
};

test(
  "a tracked or suggested paragraph deletion beside pending joins accepts as the direct one",
  async () => {
    let compared = 0;
    for (const shapes of DOCUMENTS) {
      const bytes = await createDocx(documentOf(shapes));
      for (const index of shapes.keys()) {
        const direct = await accepted(bytes, idOf(index), "direct");
        for (const mode of ["tracked-changes", "suggested"] as const) {
          const reviewed = await accepted(bytes, idOf(index), mode);
          // What an earlier deletion already took whole records nothing.
          if (reviewed === undefined) continue;
          compared += 1;
          expect({ shapes, index, mode, text: reviewed }).toEqual({
            shapes,
            index,
            mode,
            text: direct ?? "",
          });
        }
      }
    }
    expect(compared).toBeGreaterThan(DOCUMENTS.length * 3);
  },
  propertyTestTimeout(60_000),
);
