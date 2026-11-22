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
  "legacy",
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
