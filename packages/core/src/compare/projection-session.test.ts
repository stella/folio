import { expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { HeaderFooter, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { applyComparison, parseComparison, planComparison } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2026-09-11T00:00:00.000Z",
} as const;
const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical"];
const HEADER_RELATIONSHIP_ID = "rId_projection_header";
const FOOTER_RELATIONSHIP_ID = "rId_projection_footer";

const paragraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const headerFooter = (
  type: "header" | "footer",
  text: string,
  paraId: string,
): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [paragraph(text, paraId)],
});

const storyMatrixDocx = (label: string): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph(`main ${label}`, "A1000001")];
  document.package.headers = new Map([
    [
      HEADER_RELATIONSHIP_ID,
      headerFooter("header", `header ${label}`, "A1000002"),
    ],
  ]);
  document.package.footers = new Map([
    [
      FOOTER_RELATIONSHIP_ID,
      headerFooter("footer", `footer ${label}`, "A1000003"),
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
    footerReferences: [{ type: "default", rId: FOOTER_RELATIONSHIP_ID }],
  };
  document.package.footnotes = [
    {
      type: "footnote",
      id: 2,
      noteType: "normal",
      content: [paragraph(`footnote ${label}`, "A1000004")],
    },
  ];
  document.package.endnotes = [
    {
      type: "endnote",
      id: 3,
      noteType: "normal",
      content: [paragraph(`endnote ${label}`, "A1000005")],
    },
  ];
  return createDocx(document);
};

// The counters instrument the reviewed-story documents the legacy projection
// walks; the canonical session projects its stories from the model instead.

test.each(SESSIONS)(
  "%s: comparison consumes the retained story projections through apply verification",
  async (session) => {
    const parsed = await parseComparison(
      await storyMatrixDocx("before"),
      await storyMatrixDocx("after"),
      { ...OPTIONS, session },
    );
    if (parsed.isErr()) {
      throw parsed.error;
    }
    expect(parsed.value.pairs).toHaveLength(5);

    const planned = planComparison(parsed.value);
    if (planned.isErr()) {
      throw planned.error;
    }
    expect(planned.value).toHaveLength(5);
    const applied = applyComparison(parsed.value, planned.value);
    if (applied.isErr()) {
      throw applied.error;
    }

    expect(applied.value.verification).toEqual({ status: "verified" });
  },
);
