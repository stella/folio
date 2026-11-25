/**
 * A document that repeats a move range is repaired on open, on both engines:
 * the first range of its id is kept, each further one is dropped and reported.
 * A save writes the repair only into a paragraph it rewrites anyway, so an
 * unedited save stays byte-identical.
 */
import { describe, expect, test } from "bun:test";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document, ParagraphContent } from "../types/document";
import { openCanonicalDocx } from "./canonicalSessionInput";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { normalizeTrackedMoveRanges } from "./trackedMoveRangeNormalization";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const range = (id: number, text: string): ParagraphContent[] => [
  {
    type: "moveToRangeStart",
    id,
    name: `move${id}`,
    author: "A",
    date: "2026-01-01T00:00:00Z",
  },
  run(text),
  { type: "moveToRangeEnd", id },
];

const paragraphsOf = (contents: ParagraphContent[][]): Document => ({
  package: {
    document: {
      content: contents.map((content, index) => ({
        type: "paragraph",
        paraId: `6100000${index}`,
        content,
      })),
    },
  },
});

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const documentXml = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

const boundariesOf = (document: Document): string[][] =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((node) =>
          node.type === "moveToRangeStart" || node.type === "moveToRangeEnd"
            ? [`${node.type}:${node.id}`]
            : [],
        )
      : [],
  );

const HEADER_PART = "word/header9.xml";
const HEADER_ID = "rId909";

const headerRange = (text: string): string =>
  `<w:p><w:moveToRangeStart w:id="7" w:name="m" w:author="A" w:date="2026-01-01T00:00:00Z"/><w:r><w:t>${text}</w:t></w:r><w:moveToRangeEnd w:id="7"/></w:p>`;

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${headerRange("First.")}${headerRange("Second.")}</w:hdr>`;

/** `source` with a header that repeats a range the body holds once. */
const withHeader = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(source);
  const edit = async (
    path: string,
    before: string,
    added: string,
  ): Promise<void> => {
    const xml = (await zip.entry(path)?.text()) ?? "";
    const index = xml.lastIndexOf(before);
    expect(index).toBeGreaterThan(-1);
    zip.set(path, `${xml.slice(0, index)}${added}${xml.slice(index)}`);
  };
  await edit(
    "[Content_Types].xml",
    "</Types>",
    `<Override PartName="/${HEADER_PART}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
  );
  await edit(
    "word/_rels/document.xml.rels",
    "</Relationships>",
    `<Relationship Id="${HEADER_ID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9.xml"/>`,
  );
  await edit(
    "word/document.xml",
    "</w:body>",
    `<w:sectPr><w:headerReference w:type="default" r:id="${HEADER_ID}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>`,
  );
  zip.set(HEADER_PART, HEADER_XML);
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const repeated = (): Document =>
  paragraphsOf([
    range(7, "First."),
    range(7, "Second."),
    range(8, "Third."),
    range(7, "Fourth."),
  ]);

describe("a repeated move range", () => {
  test("keeps its first range and reports each dropped one, on both engines", async () => {
    const source = await createDocx(repeated());
    const legacy = await parseDocx(source, {
      preloadFonts: false,
      detectVariables: false,
    });
    const canonical = await openCanonicalDocx(source, {
      preloadFonts: false,
      detectVariables: false,
    });
    if (canonical.isErr()) throw canonical.error;
    for (const parsed of [legacy, canonical.value.document]) {
      expect(boundariesOf(parsed)).toEqual([
        ["moveToRangeStart:7", "moveToRangeEnd:7"],
        [],
        ["moveToRangeStart:8", "moveToRangeEnd:8"],
        [],
      ]);
      expect(
        parsed.parseWarnings?.filter(
          ({ code }) => code === PARSE_WARNING_CODES.droppedMoveRange,
        ),
      ).toHaveLength(2);
      expect(
        parsed.parseWarnings?.some(
          ({ code }) => code === PARSE_WARNING_CODES.unbalancedMoveRange,
        ) ?? false,
      ).toBe(false);
    }
  });

  for (const session of SESSIONS)
    test(`${session}: an unedited save keeps the document part byte-identical`, async () => {
      const source = await createDocx(repeated());
      const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
      expect(await documentXml(await reviewer.toBuffer())).toBe(
        await documentXml(source),
      );
    });

  // A header numbers its ranges apart from the body, and replays its bytes.
  for (const session of SESSIONS)
    test(`${session}: a header keeps its own first range, and its bytes unedited`, async () => {
      const source = await withHeader(
        await createDocx(paragraphsOf([range(7, "Body.")])),
      );
      const parsed = await parseDocx(source, {
        preloadFonts: false,
        detectVariables: false,
      });
      const header = parsed.package.headers?.get(HEADER_ID);
      expect(
        header?.content.map((block) =>
          block.type === "paragraph"
            ? block.content.map(({ type }) => type)
            : [],
        ),
      ).toEqual([["moveToRangeStart", "run", "moveToRangeEnd"], ["run"]]);
      expect(boundariesOf(parsed)).toEqual([
        ["moveToRangeStart:7", "moveToRangeEnd:7"],
      ]);
      const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
      const saved = await ZipArchive.read(await reviewer.toBuffer());
      expect(await saved.entry(HEADER_PART)?.text()).toBe(HEADER_XML);
    });

  for (const copyFirst of [true, false])
    test(`a copy its pending insertion owns stays, ${copyFirst ? "before" : "after"} the range it copies`, () => {
      const owned: ParagraphContent = {
        type: "insertion",
        info: { id: 1, author: "B", date: "2026-01-02T00:00:00Z" },
        content: [
          {
            type: "moveToRangeStart",
            id: 7,
            name: "move7",
            author: "A",
            copyOriginAncestors: [0],
          },
          { type: "run", content: [{ type: "text", text: "Copy." }] },
          { type: "moveToRangeEnd", id: 7, copyOriginAncestors: [0] },
        ],
      };
      const { document } = paragraphsOf([
        copyFirst ? [owned] : range(7, "First."),
        copyFirst ? range(7, "First.") : [owned],
        range(7, "Repeated."),
      ]).package;
      expect(normalizeTrackedMoveRanges({ documentBody: document })).toEqual({
        removedUnbalancedMoveRangeMarkers: 0,
        droppedMoveRanges: [{ kind: "moveTo", id: 7 }],
      });
      expect(
        document.content.map((block) =>
          block.type === "paragraph"
            ? block.content.map(({ type }) => type)
            : [],
        ),
      ).toEqual([
        copyFirst
          ? ["insertion"]
          : ["moveToRangeStart", "run", "moveToRangeEnd"],
        copyFirst
          ? ["moveToRangeStart", "run", "moveToRangeEnd"]
          : ["insertion"],
        ["run"],
      ]);
      expect(owned.content).toHaveLength(3);
    });

  test("a repeated half is still an unbalanced one", () => {
    const { document } = paragraphsOf([
      range(7, "First."),
      [run("Second."), { type: "moveToRangeEnd", id: 7 }],
    ]).package;
    expect(normalizeTrackedMoveRanges({ documentBody: document })).toEqual({
      removedUnbalancedMoveRangeMarkers: 1,
    });
    expect(boundariesOf({ package: { document } })).toEqual([
      ["moveToRangeStart:7", "moveToRangeEnd:7"],
      [],
    ]);
  });
});
