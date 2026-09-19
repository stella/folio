/**
 * An unbalanced inline range marker must not make folio refuse a document.
 *
 * Word opens a file whose `w:commentRangeStart` has no end, or whose
 * `w:moveToRangeEnd` has no start; edits produce them routinely, because
 * deleting the text between two markers leaves one of them behind. Folio's
 * model validator counts starts against ends and calls the imbalance an error,
 * so the parse boundary has to normalise the input before the validator sees
 * it. The validator's invariant stays: it is about what folio builds.
 *
 * The normalisers used to walk only `paragraph.content` while the validator
 * walked the whole inline tree, so a marker inside `w:ins`, `w:hyperlink`,
 * `w:sdt` or `w:bdo` was judged but never normalised. Position is therefore
 * the variable here, not the marker.
 *
 * Per kind, what Word does and what folio does:
 * - a comment range's unmatched half becomes a `w:commentReference` in place,
 *   so the comment stays anchored to the point it was anchored to;
 * - a move range's unmatched half is dropped, because half a move range
 *   delimits nothing.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";
import { validateFolioDocumentModel } from "./modelValidation";
import { createEmptyDocument } from "../utils/createDocument";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const AUTHORED = 'w:author="A" w:date="2024-01-01T00:00:00Z"';

const MARKERS = {
  commentRange: (id: number) => ({
    start: `<w:commentRangeStart w:id="${id}"/>`,
    end: `<w:commentRangeEnd w:id="${id}"/>`,
  }),
  moveFromRange: (id: number) => ({
    start: `<w:moveFromRangeStart w:id="${id}" w:name="mv${id}" ${AUTHORED}/>`,
    end: `<w:moveFromRangeEnd w:id="${id}"/>`,
  }),
  moveToRange: (id: number) => ({
    start: `<w:moveToRangeStart w:id="${id}" w:name="mv${id}" ${AUTHORED}/>`,
    end: `<w:moveToRangeEnd w:id="${id}"/>`,
  }),
} as const;

type MarkerKind = keyof typeof MARKERS;

const MARKER_KINDS = Object.keys(MARKERS) as MarkerKind[];

/** Which halves of the pair the document carries. */
const BALANCES = ["both", "startOnly", "endOnly", "twoStarts", "twoEnds"] as const;

type Balance = (typeof BALANCES)[number];

const halvesFor = (balance: Balance, marker: { start: string; end: string }): string[] => {
  switch (balance) {
    case "both":
      return [marker.start, marker.end];
    case "startOnly":
      return [marker.start];
    case "endOnly":
      return [marker.end];
    case "twoStarts":
      return [marker.start, marker.start];
    case "twoEnds":
      return [marker.end, marker.end];
    default: {
      const unreachable: never = balance;
      return unreachable;
    }
  }
};

const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;

/** The inline wrapper the markers sit inside, each transparent to a range. */
const HOSTS = {
  /** Directly in the paragraph. */
  paragraph: (inner: string) => `<w:p>${inner}</w:p>`,
  /** Inside a tracked insertion. */
  insertion: (inner: string) => `<w:p><w:ins w:id="800" ${AUTHORED}>${inner}</w:ins></w:p>`,
  /** Inside a bidirectional override. */
  bidiOverride: (inner: string) => `<w:p><w:bdo w:val="rtl">${inner}</w:bdo></w:p>`,
  /** Inside a bidirectional embedding nested in an insertion. */
  bidiInInsertion: (inner: string) =>
    `<w:p><w:ins w:id="801" ${AUTHORED}><w:dir w:val="ltr">${inner}</w:dir></w:ins></w:p>`,
  /** Inside an inline content control. */
  inlineSdt: (inner: string) =>
    `<w:p><w:sdt><w:sdtPr/><w:sdtContent>${inner}</w:sdtContent></w:sdt></w:p>`,
  /** In a paragraph inside a table cell. */
  tableCell: (inner: string) =>
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr/><w:p>${inner}</w:p></w:tc></w:tr></w:tbl>`,
} as const;

type Host = keyof typeof HOSTS;

const HOST_NAMES = Object.keys(HOSTS) as Host[];

/**
 * A package that already carries `word/comments.xml`, so a comment range in
 * the fixture names a comment that exists: a marker whose comment is missing
 * is a dangling reference, which is a different normalisation from an
 * unmatched half and would mask it.
 */
let commentBearingPackage: Promise<ArrayBuffer> | undefined;

const COMMENT_IDS = [0, 1, 2] as const;

const basePackage = (): Promise<ArrayBuffer> => {
  commentBearingPackage ??= createDocx({
    ...createEmptyDocument(),
    package: {
      ...createEmptyDocument().package,
      document: {
        ...createEmptyDocument().package.document,
        comments: COMMENT_IDS.map((id) => ({
          id,
          author: "A",
          initials: "A",
          date: "2024-01-01T00:00:00Z",
          content: [{ type: "paragraph", content: [] }],
        })),
        content: [
          {
            type: "paragraph",
            paraId: "C0000001",
            content: COMMENT_IDS.flatMap((id) => [
              { type: "commentRangeStart" as const, id },
              { type: "run" as const, content: [{ type: "text" as const, text: "seed" }] },
              { type: "commentRangeEnd" as const, id },
              { type: "commentReference" as const, id },
            ]),
          },
        ],
      },
    },
  });
  return commentBearingPackage;
};

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await basePackage());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const savedDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

type Outcome = {
  warnings: readonly string[];
  saved: string;
  savedIsValid: boolean;
};

const openAndSave = async (body: string): Promise<Outcome> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  const reopened = await parseDocx(saved, { preloadFonts: false });
  return {
    warnings: parsed.warnings ?? [],
    saved: await savedDocumentXml(saved),
    savedIsValid: validateFolioDocumentModel(reopened).valid,
  };
};

const occurrences = (xml: string, needle: string): number => xml.split(needle).length - 1;

describe("an unbalanced inline range marker never refuses the document", () => {
  test("every marker kind, in every wrapper, in every balance, opens and re-saves", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...MARKER_KINDS),
        fc.constantFrom(...HOST_NAMES),
        fc.constantFrom(...BALANCES),
        fc.constantFrom(...COMMENT_IDS),
        async (kind, host, balance, id) => {
          const marker = MARKERS[kind](id);
          const halves = halvesFor(balance, marker);
          const inner = [halves[0] ?? "", run("text"), ...halves.slice(1)].join("");

          const { saved, savedIsValid } = await openAndSave(HOSTS[host](inner));

          // What folio writes back is a package folio would accept: the
          // validator's invariant holds over the model folio built.
          expect({ kind, host, balance, valid: savedIsValid }).toEqual({
            kind,
            host,
            balance,
            valid: true,
          });

          const starts = occurrences(saved, marker.start);
          const ends = occurrences(saved, marker.end);
          if (balance === "both") {
            // A balanced pair is untouched, wherever it sits.
            expect({ kind, host, starts, ends }).toEqual({ kind, host, starts: 1, ends: 1 });
            return;
          }
          // Every unmatched half is gone, deterministically: a comment range's
          // half is rewritten as a point reference, a move range's is dropped.
          expect({ kind, host, balance, starts, ends }).toEqual({
            kind,
            host,
            balance,
            starts: 0,
            ends: 0,
          });
          if (kind === "commentRange") {
            expect(occurrences(saved, `<w:commentReference w:id="${id}"/>`)).toBe(halves.length);
          }
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
  });

  test("a lone comment range start inside a bidirectional override opens", async () => {
    const { warnings, savedIsValid } = await openAndSave(
      HOSTS.bidiOverride(`<w:commentRangeStart w:id="1"/>${run("text")}`),
    );

    expect(savedIsValid).toBe(true);
    expect(warnings.join("\n")).toContain("unbalanced comment range marker");
  });

  test("a lone move-to range end inside a tracked insertion opens", async () => {
    const { warnings, savedIsValid } = await openAndSave(
      HOSTS.bidiInInsertion(`${run("text")}<w:moveToRangeEnd w:id="42"/>`),
    );

    expect(savedIsValid).toBe(true);
    expect(warnings.join("\n")).toContain("unbalanced tracked move range marker");
  });
});
