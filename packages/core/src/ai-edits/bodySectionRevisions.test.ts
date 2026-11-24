import { expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import type { BlockContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const paragraph = (paraId: string, text: string): BlockContent => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A body whose first section ends in a paragraph carrying a tracked section change. */
const sectionDocx = () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("10000001", "Opening"),
    {
      ...paragraph("10000002", "Section end"),
      sectionProperties: {
        sectionStart: "nextPage",
        propertyChanges: [
          {
            type: "sectionPropertyChange",
            info: { id: 7, author: "Reviewer", date: "2026-11-20T00:00:00Z" },
            previousProperties: { sectionStart: "continuous" },
          },
        ],
      },
    },
    paragraph("10000003", "Closing"),
  ];
  return createDocx(document);
};

for (const session of ["canonical"] as const)
  for (const decision of ["accept", "reject"] as const)
    test(`${decision}ing every change resolves a body section's change (${session})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await sectionDocx(), {
        session,
      });
      expect(reviewer.getChanges().map(({ id }) => id)).toContain(7);
      if (decision === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      expect(reviewer.getChanges()).toEqual([]);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      const end = reopened.toDocument().package.document.content.at(1);
      if (end?.type !== "paragraph") throw new Error("expected a paragraph");
      expect(end.sectionProperties?.propertyChanges).toBeUndefined();
      expect(end.sectionProperties?.sectionStart).toBe(
        decision === "accept" ? "nextPage" : "continuous",
      );
    });
