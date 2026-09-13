import { expect, test } from "bun:test";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import type { Run } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const makeDocument = async (href: string | undefined) => {
  const document = createEmptyDocument();
  const run: Run = { type: "run", content: [{ type: "text", text: "Shared linked text" }] };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      content: href === undefined ? [run] : [{ type: "hyperlink", href, children: [run] }],
    },
  ];
  return createDocx(document);
};

for (const [baseHref, targetHref] of [
  [undefined, "https://example.test/target"],
  [undefined, "https://example.test"],
  ["https://example.test/base", undefined],
  ["https://example.test/base", "https://example.test/target"],
  [undefined, ""],
] as const) {
  test(`saved review preserves both hyperlink endpoints: ${String(baseHref)} -> ${String(targetHref)}`, async () => {
    const base = await makeDocument(baseHref);
    const target = await makeDocument(targetHref);
    const result = await compareDocx(base, target, {
      author: "Reviewer",
      timestamp: "2026-09-13T00:00:00.000Z",
    });
    if (result.isErr()) throw result.error;
    expect(result.value.verification.status).toBe("verified");
    for (const view of ["accept", "reject"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
      if (view === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      const saved = await reviewer.toBuffer();
      const document = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
      const paragraph = document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
      const links = paragraph.content.filter((item) => item.type === "hyperlink");
      const expected = view === "accept" ? targetHref : baseHref;
      expect(links.map(({ href }) => href ?? "")).toEqual(expected === undefined ? [] : [expected]);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(reopened.getChanges()).toHaveLength(0);
    }
  });
}
