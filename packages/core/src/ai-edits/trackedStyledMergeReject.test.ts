/**
 * A text box sits in the paragraph as a zero-width anchor whose marks are
 * rendered in the paragraph's style cascade. Restyling the paragraph rebases
 * the text runs but not the anchor, so the save read the anchor's old look as
 * direct formatting against the new style: a tracked merge naming Heading2
 * wrote `bold: false, fontSize: 22` into the text box's run with no change
 * recording it, and rejecting after a save kept it. The anchor carries the
 * run's own properties, which a restyle does not change.
 */

import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document, Paragraph, Run } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const BOX_HOST = "The box beside this clause restates it. ";

const buildDocument = (): Document => {
  const document = fromMarkdown(
    "# Agreement\n\nDelivery follows the schedule.",
  );
  document.package.document.content.push(
    paragraph([
      run(BOX_HOST),
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
              textBody: { content: [paragraph("Boxed note.")] },
            },
          },
        ],
      },
    ]),
  );
  return document;
};

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Test" });

/** The run formatting of the paragraph that holds the text box. */
const hostRunFormatting = (document: Document) => {
  const host = document.package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" &&
      block.content.some(
        (item) =>
          item.type === "run" &&
          item.content.some((content) => content.type === "shape"),
      ),
  );
  if (!host) throw new Error("the text box paragraph must survive");
  return host.content
    .filter((item): item is Run => item.type === "run")
    .map((item) => ({
      shape: item.content.some((content) => content.type === "shape"),
      formatting: item.formatting ?? null,
    }));
};

const openWithIds = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(buildDocument()),
    { author: "Editor" },
  );
  const idOf = (text: string) => {
    const id = reviewer
      .getContent()
      .find((block) => block.text.startsWith(text))?.id;
    if (id === undefined) throw new Error(`fixture must expose "${text}"`);
    return id;
  };
  return { reviewer, idOf };
};

describe("a restyled paragraph holding a text box", () => {
  test("a tracked merge naming a style rejects to the same document after a save", async () => {
    const { reviewer, idOf } = await openWithIds();
    const original = hostRunFormatting(reviewer.toDocument());
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: idOf("Delivery follows the schedule."),
          separator: " ",
          mergedParagraphProperties: { styleId: "Heading2" },
        },
      ],
    });
    expect(result.skipped).toEqual([]);

    const reopened = await reopen(reviewer);
    expect(hostRunFormatting(reopened.toResolvedDocument("reject"))).toEqual(
      original,
    );
    expect(hostRunFormatting(reviewer.toResolvedDocument("reject"))).toEqual(
      original,
    );
  });

  test("a direct restyle writes no old look into the text box's run", async () => {
    const { reviewer, idOf } = await openWithIds();
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "restyle",
          type: "setBlockParagraphProperties",
          blockId: idOf(BOX_HOST),
          properties: { styleId: "Heading2" },
        },
      ],
    });
    expect(result.skipped).toEqual([]);

    const box = hostRunFormatting((await reopen(reviewer)).toDocument()).find(
      ({ shape }) => shape,
    );
    expect(box).toEqual({ shape: true, formatting: null });
  });
});
