import { describe, expect, test } from "bun:test";

import type { Hyperlink, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { collectHyperlinksWithoutRId } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const NEW_HYPERLINK: Hyperlink = {
  type: "hyperlink",
  href: "https://example.invalid/new",
  children: [{ type: "run", content: [{ type: "text", text: "new" }] }],
};

const nestedTrackedHyperlinkParagraph = (): Paragraph => ({
  type: "paragraph",
  content: [
    {
      type: "inlineSdt",
      properties: { sdtType: "richText" },
      content: [
        {
          type: "insertion",
          info: { id: 1, author: "compare", date: "2026-09-12T00:00:00.000Z" },
          content: [NEW_HYPERLINK],
        },
      ],
    },
  ],
});

describe("tracked hyperlink package resources", () => {
  test("discovers new relationships through every recursive inline owner", () => {
    expect(collectHyperlinksWithoutRId([nestedTrackedHyperlinkParagraph()])).toEqual([
      NEW_HYPERLINK,
    ]);
  });

  test("selective save delegates relationship allocation to full repack", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [nestedTrackedHyperlinkParagraph()];
    const result = await attemptSelectiveSave(document, new ArrayBuffer(0), {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).toBeNull();
  });
});
