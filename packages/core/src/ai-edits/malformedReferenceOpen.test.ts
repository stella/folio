/**
 * A part holding a reference XML refuses (an entity no declaration defines, a
 * character XML 1.0 does not allow) is not XML. Both engines refuse the
 * package with one typed error naming the part, the place and the reference.
 */
import { describe, expect, test } from "bun:test";

import { MalformedXmlReferenceError } from "@stll/docx-core/xml";
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
  "an entity in the document's text": {
    part: "word/document.xml",
    reference: "nbsp",
    kind: "undeclaredEntity",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&nbsp;words."),
  },
  "an entity in a document attribute": {
    part: "word/document.xml",
    reference: "custom",
    kind: "undeclaredEntity",
    edit: (xml: string) => xml.replace("<w:p>", '<w:p w:rsidR="a&custom;">'),
  },
  "an entity in a styles attribute": {
    part: "word/styles.xml",
    reference: "copy",
    kind: "undeclaredEntity",
    edit: (xml: string) =>
      xml.replace('w:styleId="Normal"', 'w:styleId="Normal&copy;"'),
  },
  "U+0000 in the document's text": {
    part: "word/document.xml",
    reference: "#0",
    kind: "invalidCharacterReference",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&#0;words."),
  },
  "a control character in the document's text": {
    part: "word/document.xml",
    reference: "#x1",
    kind: "invalidCharacterReference",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&#x1;words."),
  },
  "a hexadecimal reference marked with a capital X": {
    part: "word/document.xml",
    reference: "#X41",
    kind: "invalidCharacterReference",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&#X41;words."),
  },
  "a decimal reference with letters among its digits": {
    part: "word/document.xml",
    reference: "#12ab",
    kind: "invalidCharacterReference",
    edit: (xml: string) => xml.replace("Plain words.", "Plain&#12ab;words."),
  },
  "a noncharacter in a styles attribute": {
    part: "word/styles.xml",
    reference: "#xFFFE",
    kind: "invalidCharacterReference",
    edit: (xml: string) =>
      xml.replace('w:styleId="Normal"', 'w:styleId="Normal&#xFFFE;"'),
  },
} as const;

describe("opening a part with a reference XML refuses", () => {
  for (const session of SESSIONS)
    for (const [name, { part, reference, kind, edit }] of Object.entries(CASES))
      test(`${session}: ${name} refuses the package with its place`, async () => {
        const { bytes, xml } = await withPart({ part, edit });
        const offset = xml.indexOf(`&${reference};`);
        expect(offset).toBeGreaterThan(-1);
        const opened = FolioDocxReviewer.fromBuffer(bytes, { session });
        await expect(opened).rejects.toBeInstanceOf(MalformedXmlReferenceError);
        await expect(opened).rejects.toMatchObject({
          part,
          offset,
          reference,
          kind,
        });
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
