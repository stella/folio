/**
 * A part naming an entity no declaration defines is not XML. Both engines
 * refuse the package with one typed error naming the part and the place.
 */
import { describe, expect, test } from "bun:test";

import { UndeclaredXmlEntityError } from "@stll/docx-core/xml";
import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = ["legacy", "canonical"] as const;

type PartEdit = { part: string; edit: (xml: string) => string };

const withPart = async ({ part, edit }: PartEdit) => {
  const zip = await ZipArchive.read(
    await createDocx(fromMarkdown("Plain words.")),
  );
  const xml = edit((await zip.entry(part)?.text()) ?? "");
  zip.set(part, xml);
  return { bytes: (await zip.generate()).slice().buffer, xml };
};

const CASES = {
  "text in the document": {
    part: "word/document.xml",
    entity: "nbsp",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&nbsp;words."),
  },
  "an attribute in the document": {
    part: "word/document.xml",
    entity: "custom",
    edit: (xml: string) => xml.replace("<w:p>", '<w:p w:rsidR="a&custom;">'),
  },
  "an attribute in the styles": {
    part: "word/styles.xml",
    entity: "copy",
    edit: (xml: string) =>
      xml.replace('w:styleId="Normal"', 'w:styleId="Normal&copy;"'),
  },
} as const;

describe("opening a part with an undeclared entity", () => {
  for (const session of SESSIONS)
    for (const [name, { part, entity, edit }] of Object.entries(CASES))
      test(`${session}: ${name} refuses the package with its place`, async () => {
        const { bytes, xml } = await withPart({ part, edit });
        const offset = xml.indexOf(`&${entity};`);
        expect(offset).toBeGreaterThan(-1);
        const opened = FolioDocxReviewer.fromBuffer(bytes, { session });
        await expect(opened).rejects.toBeInstanceOf(UndeclaredXmlEntityError);
        await expect(opened).rejects.toMatchObject({ part, offset, entity });
      });

  for (const session of SESSIONS)
    test(`${session}: built-in entities and character references still open`, async () => {
      const { bytes } = await withPart({
        part: "word/document.xml",
        edit: (xml) => xml.replace("Plain words.", "Plain&amp;&#65;&#x42;"),
      });
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
        "Plain&AB",
      ]);
    });
});
