/**
 * Rejecting a paragraph-property change restores the properties it recorded,
 * stated values included. A list item whose level supplies left 720 and a
 * 360 hanging indent was, before the change, a plain paragraph stating left
 * 720 itself: rejected, it states that indentation again, though the value
 * equals the one its numbering gave it.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = ["canonical"] as const;
const LEVEL_INDENT = 720;
const CHANGE_ID = 30;

const numberedWithStatedPrevious = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 0,
        levels: [
          {
            ilvl: 0,
            numFmt: "decimal",
            lvlText: "%1.",
            pPr: {
              indentLeft: LEVEL_INDENT,
              indentFirstLine: -360,
              hangingIndent: true,
            },
          },
        ],
      },
    ],
    nums: [{ numId: 1, abstractNumId: 0 }],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "10000002",
      formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: {
            id: CHANGE_ID,
            author: "Reviewer",
            date: "2026-11-21T00:00:00Z",
          },
          previousFormatting: { indentLeft: LEVEL_INDENT },
        },
      ],
      content: [{ type: "run", content: [{ type: "text", text: "Beta" }] }],
    },
  ];
  return await createDocx(document);
};

/** The first paragraph's `w:pPr` as saved, or "" when it states none. */
const savedProperties = async (bytes: ArrayBuffer): Promise<string> => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const paragraph = xml.match(/<w:p\b.*?<\/w:p>/su)?.[0] ?? "";
  return (
    paragraph.match(/^<w:p\b[^>]*>(<w:pPr>.*?<\/w:pPr>|<w:pPr\/>)/su)?.[1] ?? ""
  );
};

const STATED = `<w:pPr><w:ind w:left="${String(LEVEL_INDENT)}"/></w:pPr>`;

describe("rejecting a property change that removed a stated indentation", () => {
  test.each(SESSIONS)(
    "%s: the restored paragraph states its indentation",
    async (session) => {
      for (const resolve of ["all", "single"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await numberedWithStatedPrevious(),
          { session },
        );
        if (resolve === "all") {
          expect(reviewer.rejectAll()).toMatchObject({ resolved: 1 });
        } else {
          const change =
            reviewer.getChanges().find(({ id }) => id === CHANGE_ID) ??
            panic("The property change is listed.");
          expect(reviewer.rejectChange(change)).toBe(true);
        }
        const saved = await reviewer.toBuffer();
        expect(await savedProperties(saved)).toBe(STATED);

        const reopened = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        expect(await savedProperties(await reopened.toBuffer())).toBe(STATED);
      }
    },
  );
});
