import { expect, test } from "bun:test";

import { OP_STORIES, paragraphsThroughTextBoxes } from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import {
  holdsUniqueParagraphIds,
  prepareCanonicalDocumentInput,
} from "./canonicalSessionInput";

/** A body paragraph drawing a text box whose paragraph holds `boxParaId`. */
const boxedDocument = (boxParaId: string | undefined): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            { type: "run", content: [{ type: "text", text: "Host " }] },
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    id: "box",
                    size: { width: 2_000_000, height: 800_000 },
                    textBody: {
                      content: [
                        {
                          type: "paragraph",
                          ...(boxParaId !== undefined && { paraId: boxParaId }),
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "Boxed note." }],
                            },
                          ],
                        },
                      ],
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

// A canonical session addresses a text box's paragraphs by their ids, as it
// does the story's own: a host model whose box paragraph holds none opens
// through its package, which gives it one.
test("a host model's text box paragraph without an id opens with one", async () => {
  expect(holdsUniqueParagraphIds(boxedDocument("34567890"))).toBe(true);
  const unnamed = boxedDocument(undefined);
  expect(holdsUniqueParagraphIds(unnamed)).toBe(false);
  const prepared = (await prepareCanonicalDocumentInput(unnamed)).unwrap();
  const ids = paragraphsThroughTextBoxes(prepared, OP_STORIES.MAIN).map(
    ({ paragraph }) => paragraph.paraId,
  );
  expect(ids).toHaveLength(2);
  for (const id of ids) expect(id).toMatch(/^[0-9A-F]{8}$/u);
});
