import { describe, expect, test } from "bun:test";

import { hashFolioAIBlockText } from "./snapshot";
import { getFolioDocumentOutline, readFolioDocumentSection } from "./scoped-reading";
import type { FolioAIBlock, FolioAIEditSnapshot } from "./types";

const makeSnapshot = (blocks: FolioAIBlock[]): FolioAIEditSnapshot => ({
  blocks,
  anchors: Object.fromEntries(
    blocks.map((block, index) => [
      block.id,
      {
        id: block.id,
        from: index * 10,
        to: index * 10 + block.text.length,
        text: block.text,
        normalizedText: block.text,
        textHash: hashFolioAIBlockText(block.text),
        hashOccurrenceCount: 1,
      },
    ]),
  ),
});

const snapshot = makeSnapshot([
  { id: "h1", kind: "heading", headingLevel: 1, text: "Agreement" },
  { id: "p1", kind: "paragraph", text: "Opening" },
  { id: "h2", kind: "heading", headingLevel: 2, text: "Payment" },
  { id: "p2", kind: "paragraph", text: "Net 30" },
  { id: "h3", kind: "heading", headingLevel: 2, text: "Termination" },
  { id: "p3", kind: "paragraph", text: "Thirty days" },
  { id: "h4", kind: "heading", headingLevel: 1, text: "Schedules" },
]);

describe("scoped document reading", () => {
  test("builds ordered section handles with explicit hierarchy", () => {
    const outline = getFolioDocumentOutline(snapshot);

    expect(outline.sections.map(({ headingBlockId, level }) => [headingBlockId, level])).toEqual([
      ["h1", 1],
      ["h2", 2],
      ["h3", 2],
      ["h4", 1],
    ]);
    expect(outline.sections.at(1)?.parentHandle).toEqual(outline.sections.at(0)?.handle);
    expect(outline.sections.at(2)?.parentHandle).toEqual(outline.sections.at(0)?.handle);
    expect(outline.sections.at(3)?.parentHandle).toBeUndefined();
  });

  test("reads through nested headings until the next peer heading", () => {
    const outline = getFolioDocumentOutline(snapshot);
    const agreement = outline.sections.at(0);
    const payment = outline.sections.at(1);
    if (agreement === undefined || payment === undefined) {
      throw new Error("Expected outline entries");
    }

    const agreementRead = readFolioDocumentSection(snapshot, agreement.handle);
    const paymentRead = readFolioDocumentSection(snapshot, payment.handle);

    expect(agreementRead.status).toBe("found");
    if (agreementRead.status === "found") {
      expect(agreementRead.section.blocks.map(({ id }) => id)).toEqual([
        "h1",
        "p1",
        "h2",
        "p2",
        "h3",
        "p3",
      ]);
    }
    expect(paymentRead.status).toBe("found");
    if (paymentRead.status === "found") {
      expect(paymentRead.section.blocks.map(({ id }) => id)).toEqual(["h2", "p2"]);
    }
  });

  test("distinguishes stale or structurally changed handles from deleted headings", () => {
    const handle = getFolioDocumentOutline(snapshot).sections.at(0)?.handle;
    if (handle === undefined) {
      throw new Error("Expected outline handle");
    }
    const renamed = makeSnapshot([
      { id: "h1", kind: "heading", headingLevel: 1, text: "Renamed agreement" },
    ]);
    const relevelled = makeSnapshot([
      { id: "h1", kind: "heading", headingLevel: 2, text: "Agreement" },
    ]);

    expect(readFolioDocumentSection(renamed, handle)).toEqual({ status: "stale" });
    expect(readFolioDocumentSection(relevelled, handle)).toEqual({ status: "stale" });
    expect(readFolioDocumentSection(makeSnapshot([]), handle)).toEqual({ status: "missing" });
  });
});
