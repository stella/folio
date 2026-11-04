/**
 * A move source keeps its run text as `w:t`.
 *
 * ISO/IEC 29500-1 §17.3.3.7 reserves `w:delText` for runs in a `w:del` region;
 * a `w:moveFrom` (§17.13.5.22) is not one, and the spec's own example spells
 * its run text `w:t`. Every host that reaches the move source writer is
 * covered here, from both spellings a reader can meet: the spec's `w:t` and
 * the `w:delText` some writers put there.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseParagraph } from "./paragraphParser";
import { createEmptyDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MOVE = 'w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z"';
const INSERTION = 'w:id="4" w:author="B" w:date="2026-01-01T00:00:00Z"';

type Spelling = { text: "t" | "delText"; instruction: string };

const SPELLINGS = {
  spec: { text: "t", instruction: "instrText" },
  deletedText: { text: "delText", instruction: "delInstrText" },
} as const satisfies Record<string, Spelling>;

const run = ({ text }: Spelling, value: string): string =>
  `<w:r><w:${text} xml:space="preserve">${value}</w:${text}></w:r>`;

const moveFrom = (content: string): string =>
  `<w:moveFrom ${MOVE}>${content}</w:moveFrom>`;

/** Each host that reaches the move source writer, written in one spelling. */
const HOSTS = {
  paragraph: (s: Spelling) => moveFrom(run(s, "moved away")),
  "inline content control": (s: Spelling) =>
    `<w:sdt><w:sdtPr/><w:sdtContent>${moveFrom(run(s, "moved away"))}</w:sdtContent></w:sdt>`,
  "tracked insertion": (s: Spelling) =>
    `<w:ins ${INSERTION}>${moveFrom(run(s, "moved away"))}</w:ins>`,
  "transparent wrapper": (s: Spelling) =>
    moveFrom(`<w:dir w:val="rtl">${run(s, "moved away")}</w:dir>`),
  hyperlink: (s: Spelling) =>
    moveFrom(
      `<w:hyperlink w:anchor="target">${run(s, "moved away")}</w:hyperlink>`,
    ),
  "simple field": (s: Spelling) =>
    moveFrom(
      `<w:fldSimple w:instr=" PAGE ">${run(s, "moved away")}</w:fldSimple>`,
    ),
  "complex field": (s: Spelling) =>
    moveFrom(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        `<w:r><w:${s.instruction} xml:space="preserve"> PAGE </w:${s.instruction}></w:r>` +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        run(s, "moved away") +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    ),
} as const satisfies Record<string, (spelling: Spelling) => string>;

const parse = (inner: string) => {
  const node = parseXmlDocument(`<w:p ${NS}>${inner}</w:p>`);
  if (!node) throw new TypeError("the fixture did not parse");
  return parseParagraph(node, new Map(), null, null);
};

const save = (inner: string): string => serializeParagraph(parse(inner));

const bodyOf = (paragraphXml: string): string =>
  paragraphXml.replace(/^<w:p>/u, "").replace(/<\/w:p>$/u, "");

const moveSources = (xml: string): string[] => {
  const sources = [
    ...xml.matchAll(/<w:moveFrom\b[^>]*>([\s\S]*?)<\/w:moveFrom>/gu),
  ].map((match) => match[1] ?? "");
  expect(sources.length).toBeGreaterThan(0);
  return sources;
};

describe("a move source spells its run text w:t", () => {
  for (const [host, markup] of Object.entries(HOSTS)) {
    for (const [name, spelling] of Object.entries(SPELLINGS)) {
      test(`${host}, read from ${name} spelling`, () => {
        const saved = save(markup(spelling));
        const sources = moveSources(saved).join("");

        expect(sources).not.toMatch(/<w:del(?:Text|InstrText)\b/u);
        expect(sources).toMatch(/<w:t(?:\s[^>]*)?>moved away<\/w:t>/u);
        if (host === "complex field") {
          expect(sources).toMatch(
            /<w:instrText(?:\s[^>]*)?> PAGE <\/w:instrText>/u,
          );
        }
        // A second save changes nothing.
        expect(save(bodyOf(saved))).toBe(saved);
      });
    }

    test(`${host}: both spellings read as the same model and save alike`, () => {
      expect(parse(markup(SPELLINGS.deletedText))).toEqual(
        parse(markup(SPELLINGS.spec)),
      );
      expect(save(markup(SPELLINGS.deletedText))).toBe(
        save(markup(SPELLINGS.spec)),
      );
    });
  }

  test("a deletion still spells its run text w:delText", () => {
    const saved = save(
      `<w:del ${MOVE}>${run(SPELLINGS.spec, "gone")}</w:del>` +
        `<w:sdt><w:sdtPr/><w:sdtContent><w:del ${INSERTION}>${run(SPELLINGS.spec, "gone")}</w:del></w:sdtContent></w:sdt>`,
    );
    expect(saved).not.toMatch(/<w:t(?:\s[^>]*)?>gone<\/w:t>/u);
    expect(
      saved.match(/<w:delText(?:\s[^>]*)?>gone<\/w:delText>/gu),
    ).toHaveLength(2);
  });
});

const packageWithBody = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  const documentXml = await zip.entry("word/document.xml")?.text();
  if (!documentXml) throw new TypeError("the package has no document part");
  zip.set(
    "word/document.xml",
    documentXml.replace(
      /<w:body>[\s\S]*<\/w:body>/u,
      `<w:body>${body}<w:sectPr/></w:body>`,
    ),
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const documentPart = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await ZipArchive.read(buffer)).entry("word/document.xml");
  if (!part) throw new TypeError("the package has no document part");
  return part.text();
};

describe("a move saved in a package", () => {
  // The spec's own move: a named source range holding `w:t`, and its
  // destination.
  const MOVE_BODY =
    "<w:p>" +
    '<w:moveFromRangeStart w:id="1" w:name="move1" w:author="A" w:date="2026-01-01T00:00:00Z"/>' +
    `<w:moveFrom ${MOVE}><w:r><w:t>Some moved text.</w:t></w:r></w:moveFrom>` +
    '<w:moveFromRangeEnd w:id="1"/>' +
    "</w:p>" +
    "<w:p>" +
    '<w:moveToRangeStart w:id="2" w:name="move1" w:author="A" w:date="2026-01-01T00:00:00Z"/>' +
    `<w:moveTo w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Some moved text.</w:t></w:r></w:moveTo>` +
    '<w:moveToRangeEnd w:id="2"/>' +
    "</w:p>";

  test("keeps w:t in the source through save and reopen", async () => {
    const opened = await FolioDocxReviewer.fromBuffer(
      await packageWithBody(MOVE_BODY),
    );
    const saved = await opened.toBuffer();
    const savedXml = await documentPart(saved);

    expect(savedXml).not.toMatch(/<w:del(?:Text|InstrText)\b/u);
    for (const source of moveSources(savedXml)) {
      expect(source).toMatch(/<w:t(?:\s[^>]*)?>Some moved text\.<\/w:t>/u);
    }

    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(reopened.getChanges()).toEqual(opened.getChanges());
    expect(await documentPart(await reopened.toBuffer())).toBe(savedXml);
  });
});
