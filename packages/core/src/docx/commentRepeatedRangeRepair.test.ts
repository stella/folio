/**
 * A comment owns one range. A package repeating a comment's range is repaired
 * on open, on both engines: its first range stays, the repeated anchors go,
 * and the open reports the text they covered. A save writes the repair only
 * where it rewrites the package anyway, so an unedited save and the parts and
 * paragraphs an edit leaves alone keep their bytes.
 */
import { panic, type Result } from "better-result";
import { describe, expect, test } from "bun:test";

import { type Document, PARSE_WARNING_CODES } from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  storiesDocument,
  TEXT_BOX_TEXT,
} from "../../../../test/consumer-scenarios/support/documents";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { openCanonicalDocx } from "./canonicalSessionInput";

const COMMENT_ID = "41";
const DOCUMENT_PART = "word/document.xml";
const EDITED_WORDS = "Payment is due on receipt.";

const start = `<w:commentRangeStart w:id="${COMMENT_ID}"/>`;
const end = `<w:commentRangeEnd w:id="${COMMENT_ID}"/>`;
const reference = `<w:r><w:commentReference w:id="${COMMENT_ID}"/></w:r>`;

/** Markers around the run reading `text`, wherever the package holds it. */
type MarkedRun = { text: string; before?: string; after?: string };

type RepeatShape = {
  name: string;
  runs: readonly MarkedRun[];
  /** The first range's text, as a review reads it. */
  kept: string;
  /** The text of each range the repair drops. */
  dropped: readonly string[];
  /** The part holding the repeated anchors. */
  part: RegExp;
};

const FIRST: MarkedRun = {
  text: "The parties agree as follows.",
  before: start,
  after: `${end}${reference}`,
};
const FIRST_TEXT = "The parties agree as follows.";
const BODY = /^word\/document\.xml$/u;

const repeatOf = (text: string): MarkedRun => ({
  text,
  before: start,
  after: end,
});

const SHAPES: readonly RepeatShape[] = [
  {
    name: "in one paragraph",
    runs: [
      {
        text: "Delivery follows the schedule",
        before: start,
        after: `${end}${reference}`,
      },
      repeatOf("warranty"),
    ],
    kept: "Delivery follows the schedule",
    dropped: ["warranty"],
    part: BODY,
  },
  {
    name: "in another paragraph, twice",
    runs: [FIRST, repeatOf("warranty"), repeatOf("Prices exclude taxes.")],
    kept: FIRST_TEXT,
    dropped: ["warranty", "Prices exclude taxes."],
    part: BODY,
  },
  {
    name: "across paragraphs",
    runs: [
      FIRST,
      { text: "warranty", before: start },
      { text: "The box beside this clause restates it. ", after: end },
    ],
    kept: FIRST_TEXT,
    dropped: ["warranty applies.\nThe box beside this clause restates it. "],
    part: BODY,
  },
  {
    name: "in a table cell",
    runs: [FIRST, repeatOf("Gadget")],
    kept: FIRST_TEXT,
    dropped: ["Gadget"],
    part: BODY,
  },
  {
    name: "in a header",
    runs: [FIRST, repeatOf("Master Agreement header")],
    kept: FIRST_TEXT,
    dropped: ["Master Agreement header"],
    part: /^word\/header\d+\.xml$/u,
  },
  {
    name: "in a footer",
    runs: [FIRST, repeatOf("Even page ")],
    kept: FIRST_TEXT,
    dropped: ["Even page "],
    part: /^word\/footer\d+\.xml$/u,
  },
  {
    name: "in a footnote",
    runs: [FIRST, repeatOf("Schedules are part of this agreement.")],
    kept: FIRST_TEXT,
    dropped: ["Schedules are part of this agreement."],
    part: /^word\/footnotes\.xml$/u,
  },
  {
    name: "in an endnote",
    runs: [FIRST, repeatOf("Warranty terms are in annex two.")],
    kept: FIRST_TEXT,
    dropped: ["Warranty terms are in annex two."],
    part: /^word\/endnotes\.xml$/u,
  },
  {
    name: "in a text box",
    runs: [FIRST, repeatOf(TEXT_BOX_TEXT)],
    kept: FIRST_TEXT,
    dropped: [TEXT_BOX_TEXT],
    part: BODY,
  },
  {
    name: "in a text box its first range covers",
    runs: [
      { text: "The box beside this clause restates it. ", before: start },
      repeatOf(TEXT_BOX_TEXT),
      { text: "Notices go to ", after: `${end}${reference}` },
    ],
    kept: `The box beside this clause restates it. ${TEXT_BOX_TEXT}Notices go to `,
    dropped: [TEXT_BOX_TEXT],
    part: BODY,
  },
  {
    name: "nested in it or overlapping it",
    runs: [
      { text: "Delivery follows the schedule", before: start },
      { text: " and the ", before: start },
      { text: "warranty", after: `${end}${reference}` },
      { text: " applies", after: end },
    ],
    kept: "Delivery follows the schedule[^1] and the warranty",
    dropped: [" and the warranty applies"],
    part: BODY,
  },
  {
    name: "overlapping across paragraphs",
    runs: [
      { text: "Delivery follows the schedule", before: start },
      { text: "warranty", before: start },
      {
        text: "The box beside this clause restates it. ",
        after: `${end}${reference}`,
      },
      { text: "Notices go to ", after: end },
    ],
    kept: "Delivery follows the schedule[^1] and the warranty applies[^e1].The box beside this clause restates it. ",
    dropped: [
      "warranty applies.\nThe box beside this clause restates it. \nNotices go to ",
    ],
    part: BODY,
  },
];

/** Words beside each shape's repeat, or around its text box, and their story. */
const EDITS_OF_REPEATS = [
  { name: "in one paragraph", words: " applies", type: "main" },
  { name: "in a table cell", words: "Gadget", type: "main" },
  { name: "in a header", words: "Agreement", type: "header" },
  { name: "in a footer", words: "Even", type: "footer" },
  { name: "in a footnote", words: "Schedules", type: "footnote" },
  { name: "in an endnote", words: "Warranty", type: "endnote" },
  { name: "in a text box", words: "The box beside", type: "main" },
  {
    name: "nested in it or overlapping it",
    words: " applies",
    type: "main",
  },
] as const;

const escapeRegExp = (text: string): string =>
  text.replaceAll(/[.*+?^${}()|[\]\\]/gu, "\\$&");

const runOf = (text: string): RegExp =>
  new RegExp(
    `<w:r(?:>| [^>]*>)(?:(?!</w:r>|<w:r[ >])[\\s\\S])*?<w:t(?: [^>]*)?>${escapeRegExp(text)}</w:t></w:r>`,
    "u",
  );

const fixture = storiesDocument();

const packageOf = async ({ runs }: RepeatShape): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await fixture);
  const parts = new Map<string, string>();
  for (const path of zip.paths())
    if (/^word\/[a-z]+\d*\.xml$/u.test(path))
      parts.set(path, (await zip.entry(path)?.text()) ?? "");
  parts.set(
    DOCUMENT_PART,
    (parts.get(DOCUMENT_PART) ?? "")
      .replaceAll(/<w:commentRange(?:Start|End) w:id="41"\/>/gu, "")
      .replaceAll(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*<w:commentReference w:id="41"\/><\/w:r>/gu,
        "",
      ),
  );
  for (const { text, before = "", after = "" } of runs) {
    const pattern = runOf(text);
    const held = [...parts].find(([, xml]) => pattern.test(xml));
    if (held === undefined) panic(`no run reads "${text}"`);
    const [path, xml] = held;
    parts.set(
      path,
      xml.replace(pattern, (markup) => `${before}${markup}${after}`),
    );
  }
  for (const [path, xml] of parts) zip.set(path, new TextEncoder().encode(xml));
  return (await zip.generate({ level: 6 })).slice().buffer;
};

const unwrap = <T, E>(result: Result<T, E>): T => {
  if (result.isErr()) throw result.error;
  return result.value;
};

const options = { preloadFonts: false, detectVariables: false } as const;

const opened = async (bytes: ArrayBuffer): Promise<Document> =>
  unwrap(await openCanonicalDocx(bytes, options)).document;

type Anchors = { starts: number; ends: number; references: number };

/** How many of the comment's anchors a model holds, in every story. */
const anchorsIn = (value: unknown, anchors: Anchors): Anchors => {
  if (value instanceof Map) {
    for (const child of value.values()) anchorsIn(child, anchors);
    return anchors;
  }
  if (Array.isArray(value)) {
    for (const child of value) anchorsIn(child, anchors);
    return anchors;
  }
  if (typeof value !== "object" || value === null) return anchors;
  if (Reflect.get(value, "id") === Number(COMMENT_ID)) {
    const type: unknown = Reflect.get(value, "type");
    if (type === "commentRangeStart") anchors.starts += 1;
    if (type === "commentRangeEnd") anchors.ends += 1;
    if (type === "commentReference") anchors.references += 1;
  }
  // The comment's own record is no anchor, and sections repeat the body.
  for (const [key, child] of Object.entries(value))
    if (key !== "comments" && key !== "sections") anchorsIn(child, anchors);
  return anchors;
};

const modelAnchors = ({ package: pkg }: Document): Anchors =>
  anchorsIn(
    [pkg.document, pkg.headers, pkg.footers, pkg.footnotes, pkg.endnotes],
    { starts: 0, ends: 0, references: 0 },
  );

const storyParts = async (bytes: ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await ZipArchive.read(bytes);
  const parts = new Map<string, string>();
  for (const path of zip.paths())
    if (
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$/u.test(
        path,
      )
    )
      parts.set(path, (await zip.entry(path)?.text()) ?? "");
  return parts;
};

const paragraphsOf = (xml: string): string[] =>
  xml.split(/(?=<w:p[ >])/u).filter((markup) => markup.startsWith("<w:p"));

const countOf = (xml: string, marker: string): number =>
  xml.split(marker).length - 1;

const replaceWords = (
  reviewer: FolioDocxReviewer,
  words: string,
  replacement: string,
): void => {
  const block = reviewer
    .snapshot()
    .blocks.find(({ text }) => text.includes(words));
  if (block === undefined) panic(`no paragraph reads "${words}"`);
  const { applied, skipped } = reviewer.applyOperations(
    [
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: block.id,
        find: words,
        replace: replacement,
      },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );
  expect(skipped).toEqual([]);
  expect(applied).toHaveLength(1);
};

describe("a package repeating a comment's range", () => {
  for (const shape of SHAPES) {
    test(`${shape.name}: the first range stays, the repeats go and are reported`, async () => {
      const document = await opened(await packageOf(shape));
      expect(modelAnchors(document)).toEqual({
        starts: 1,
        ends: 1,
        references: 1,
      });
      expect(
        document.parseWarnings?.flatMap(({ code, value }) =>
          code === PARSE_WARNING_CODES.droppedCommentRange ? [value] : [],
        ),
      ).toEqual([...shape.dropped]);
      expect(
        document.parseWarnings?.filter(
          ({ code }) =>
            code === PARSE_WARNING_CODES.unpairedCommentRangeBoundary ||
            code === PARSE_WARNING_CODES.unbalancedCommentRange,
        ),
      ).toEqual([]);
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await packageOf(shape),
      );
      expect(
        reviewer.getComments().map(({ anchoredText }) => anchoredText),
      ).toEqual([shape.kept]);
    });

    test(`${shape.name}: an unedited save keeps every story's bytes`, async () => {
      const source = await packageOf(shape);
      const reviewer = await FolioDocxReviewer.fromBuffer(source);
      expect(await storyParts(await reviewer.toBuffer())).toEqual(
        await storyParts(source),
      );
    });

    test(`${shape.name}: an edit elsewhere leaves the repeated anchors as authored`, async () => {
      const source = await packageOf(shape);
      const reviewer = await FolioDocxReviewer.fromBuffer(source);
      replaceWords(reviewer, EDITED_WORDS, "Payment is due at once.");
      const before = await storyParts(source);
      const after = await storyParts(await reviewer.toBuffer());
      for (const [part, xml] of before) {
        if (part !== DOCUMENT_PART) {
          expect(after.get(part)).toBe(xml);
          continue;
        }
        const edited = new Set(paragraphsOf(after.get(part) ?? ""));
        expect(
          paragraphsOf(xml).filter((paragraph) => !edited.has(paragraph)),
        ).toEqual([expect.stringContaining(EDITED_WORDS)]);
      }
      const repeated = [...after].find(
        ([part, xml]) => shape.part.test(part) && xml.includes(start),
      );
      expect(repeated).toBeDefined();
    });

    test(`${shape.name}: a save after the open's repair reopens with the first range`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await packageOf(shape),
      );
      replaceWords(reviewer, EDITED_WORDS, "Payment is due at once.");
      const saved = await reviewer.toBuffer();
      expect(modelAnchors(await opened(saved))).toEqual({
        starts: 1,
        ends: 1,
        references: 1,
      });
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(
        reopened.getComments().map(({ anchoredText }) => anchoredText),
      ).toEqual([shape.kept]);
    });

    test(`${shape.name}: an edit of the repaired model inverts exactly`, async () => {
      const bytes = await packageOf(shape);
      const before = await opened(bytes);
      const original = (await opened(bytes)).package;
      const first = before.package.document.content.find(
        (block) =>
          block.type === "paragraph" &&
          JSON.stringify(block).includes('"commentRangeStart"'),
      );
      if (first?.type !== "paragraph" || first.paraId === undefined)
        panic("no commented paragraph");
      const ops: DocumentOp[] = [
        {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: {
            story: OP_STORIES.MAIN,
            blockId: first.paraId,
            offset: 0,
            zeroWidthBefore: 0,
          },
          to: { story: OP_STORIES.MAIN, blockId: first.paraId, offset: 3 },
        },
      ];
      const edited = unwrap(applyDocumentOps(before, ops));
      const inverse: DocumentOp[] = JSON.parse(JSON.stringify(edited.inverse));
      const undone = unwrap(applyDocumentOps(edited.document, inverse));
      expect(undone.document.package).toEqual(original);
    });
  }

  for (const { name, words, type } of EDITS_OF_REPEATS) {
    const shape = SHAPES.find((candidate) => candidate.name === name);
    if (shape === undefined) panic(`no shape ${name}`);
    test(`${name}: editing what holds the repeat writes the repair there`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await packageOf(shape),
      );
      const held = reviewer
        .listStories()
        .flatMap(({ handle }) =>
          (reviewer.snapshotStory(handle)?.blocks ?? []).map((block) => ({
            handle,
            block,
          })),
        )
        .find(
          ({ handle, block }) =>
            handle.type === type && block.text.includes(words),
        );
      if (held === undefined) panic(`no ${type} reads "${words}"`);
      const { skipped } = reviewer.applyDocumentOperationsToStory({
        story: held.handle,
        batch: {
          version: 1,
          mode: "direct",
          operations: [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId: held.block.id,
              find: words,
              replace: "Edited",
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(skipped).toEqual([]);
      const bytes = await reviewer.toBuffer();
      const saved = [...(await storyParts(bytes)).values()].join("");
      expect(countOf(saved, start)).toBe(1);
      expect(countOf(saved, end)).toBe(1);
      const reopened = await FolioDocxReviewer.fromBuffer(bytes);
      expect(
        reopened.getComments().map(({ anchoredText }) => anchoredText),
      ).toEqual([shape.kept]);
    });
  }
});
