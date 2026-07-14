import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document, InlineSdt, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

const reviewedDocument = (): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "inlineSdt",
                properties: {
                  sdtType: "richText",
                  alias: "Reviewed clause",
                  tag: "reviewed-clause",
                  id: 42,
                },
                content: [
                  { type: "run", content: [{ type: "text", text: "Clause " }] },
                  {
                    type: "deletion",
                    info: { id: 1, author: "Reviewer", date: "2026-07-14T08:00:00Z" },
                    content: [{ type: "run", content: [{ type: "text", text: "old" }] }],
                  },
                  {
                    type: "insertion",
                    info: { id: 2, author: "Reviewer", date: "2026-07-14T08:01:00Z" },
                    content: [{ type: "run", content: [{ type: "text", text: "new" }] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file("word/document.xml");
  if (!part) {
    throw new Error("missing word/document.xml");
  }
  return part.async("text");
};

const firstInlineSdt = (paragraph: Paragraph): InlineSdt => {
  const sdt = paragraph.content.find((content) => content.type === "inlineSdt");
  if (!sdt || sdt.type !== "inlineSdt") {
    throw new Error("missing inline content control");
  }
  return sdt;
};

const inlineSdtText = (sdt: InlineSdt): string =>
  sdt.content
    .flatMap((content) => {
      if (content.type === "run") {
        return content.content;
      }
      if (content.type === "insertion" || content.type === "deletion") {
        return content.content.flatMap((run) => (run.type === "run" ? run.content : []));
      }
      return [];
    })
    .map((content) => (content.type === "text" ? content.text : ""))
    .join("");

const parsedInlineSdt = async (buffer: ArrayBuffer): Promise<InlineSdt> => {
  const parsed = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  const paragraph = parsed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("missing paragraph");
  }
  return firstInlineSdt(paragraph);
};

describe("inline content-control tracked changes", () => {
  test("round-trips wrappers inside w:sdtContent", async () => {
    const buffer = await createDocx(reviewedDocument());
    const xml = await documentXml(buffer);
    const sdtContent = xml.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/u)?.at(1) ?? "";

    expect(sdtContent).toContain("<w:del ");
    expect(sdtContent).toContain("<w:ins ");

    const parsed = await parsedInlineSdt(buffer);
    expect(parsed.content.map((content) => content.type)).toEqual(["run", "deletion", "insertion"]);
  });

  test("accept and reject keep the content control with the correct text", async () => {
    const buffer = await createDocx(reviewedDocument());

    const accepting = await FolioDocxReviewer.fromBuffer(buffer);
    expect(accepting.getChanges()).toHaveLength(2);
    expect(accepting.acceptAll()).toBe(2);
    const accepted = await accepting.toBuffer();
    const acceptedSdt = await parsedInlineSdt(accepted);
    expect(inlineSdtText(acceptedSdt)).toBe("Clause new");
    expect(acceptedSdt.content.every((content) => content.type === "run")).toBe(true);

    const rejecting = await FolioDocxReviewer.fromBuffer(buffer);
    expect(rejecting.getChanges()).toHaveLength(2);
    expect(rejecting.rejectAll()).toBe(2);
    const rejected = await rejecting.toBuffer();
    const rejectedSdt = await parsedInlineSdt(rejected);
    expect(inlineSdtText(rejectedSdt)).toBe("Clause old");
    expect(rejectedSdt.content.every((content) => content.type === "run")).toBe(true);
  });
});
