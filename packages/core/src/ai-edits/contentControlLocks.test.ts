/**
 * A content control's `w:lock` holds against every document operation, in
 * every mode.
 *
 * `contentLocked` and `sdtContentLocked` forbid any text, formatting or block
 * change inside the control; `sdtLocked` and `sdtContentLocked` forbid
 * removing it, also by removing the text or paragraph that holds it. A refused
 * operation is reported as `contentControlLock` and changes nothing.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";
import type {
  FolioAIEditApplyMode,
  FolioAIEditOperation,
  FolioAITextRangeHandle,
} from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const createDocx = async (bodyXml: string): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>${bodyXml}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const LOCKS = [
  null,
  "unlocked",
  "sdtLocked",
  "contentLocked",
  "sdtContentLocked",
] as const;
type Lock = (typeof LOCKS)[number];

const MODES = [
  "direct",
  "tracked-changes",
  "suggested",
] as const satisfies readonly FolioAIEditApplyMode[];

const controlProperties = (lock: Lock): string =>
  '<w:sdtPr><w:alias w:val="Guarded"/><w:id w:val="7"/>' +
  (lock === null ? "" : `<w:lock w:val="${lock}"/>`) +
  "</w:sdtPr>";

/** What an operation does to the control. */
type Reach =
  /** Only text beside the control. */
  | "outside"
  /** Text, formatting or blocks inside it. */
  | "content"
  /** The control itself, with what holds it. */
  | "removal"
  /** The whole paragraph holding it, which may keep or change any part. */
  | "rewrite";

const refuses = (lock: Lock, reach: Reach): boolean => {
  const content = lock === "contentLocked" || lock === "sdtContentLocked";
  const control = lock === "sdtLocked" || lock === "sdtContentLocked";
  switch (reach) {
    case "outside":
      return false;
    case "content":
      return content;
    case "removal":
      return control;
    case "rewrite":
      return content || control;
    default: {
      const unhandled: never = reach;
      throw new Error(`Unhandled reach ${String(unhandled)}`);
    }
  }
};

type Block = { id: string; text: string };

const range = (
  block: Block,
  startOffset: number,
  endOffset: number,
): FolioAITextRangeHandle => ({
  type: "textRange",
  story: "main",
  blockId: block.id,
  startOffset,
  endOffset,
  selectedTextHash: hashFolioAIBlockText(
    block.text.slice(startOffset, endOffset),
  ),
});

type Case = {
  name: string;
  reach: Reach;
  /** The operation, against the block holding the control and its neighbours. */
  operation: (blocks: { holder: Block; before: Block }) => FolioAIEditOperation;
};

type Placement = {
  kind: "inline" | "block";
  bodyXml: (lock: Lock) => string;
  /** The text the control holds. */
  controlText: string;
  holderText: string;
  cases: readonly Case[];
};

const PLACEMENTS: readonly Placement[] = [
  {
    kind: "inline",
    controlText: "CTRL",
    holderText: "Before CTRL after.",
    bodyXml: (lock) =>
      "<w:p><w:r><w:t>Opening.</w:t></w:r></w:p>" +
      '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      `<w:sdt>${controlProperties(lock)}<w:sdtContent><w:r><w:t>CTRL</w:t></w:r></w:sdtContent></w:sdt>` +
      '<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>' +
      "<w:p><w:r><w:t>Closing.</w:t></w:r></w:p>",
    cases: [
      {
        name: "replacing text beside it",
        reach: "outside",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceInBlock",
          blockId: holder.id,
          find: "Before",
          replace: "Earlier",
        }),
      },
      {
        name: "replacing its text",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceInBlock",
          blockId: holder.id,
          find: "CTRL",
          replace: "Other",
        }),
      },
      {
        name: "replacing a range inside it",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceRange",
          range: range(holder, 8, 10),
          replace: "xx",
        }),
      },
      {
        name: "formatting its text",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "formatRange",
          range: range(holder, 7, 11),
          formatting: { bold: true },
        }),
      },
      {
        name: "splitting the paragraph inside it",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "splitBlock",
          blockId: holder.id,
          offset: 9,
        }),
      },
      {
        // The replacement rewrites the control's content; the control stays.
        name: "replacing text across all of it",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceInBlock",
          blockId: holder.id,
          find: "Before CTRL after",
          replace: "Gone",
        }),
      },
      {
        name: "deleting the paragraph holding it",
        reach: "removal",
        operation: ({ holder }) => ({
          id: "op",
          type: "deleteBlock",
          blockId: holder.id,
        }),
      },
      {
        name: "merging the paragraph before into it",
        reach: "outside",
        operation: ({ before }) => ({
          id: "op",
          type: "mergeBlockWithNext",
          blockId: before.id,
        }),
      },
      {
        name: "rewriting the paragraph holding it",
        reach: "rewrite",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceBlock",
          blockId: holder.id,
          text: "New words.",
        }),
      },
      {
        name: "inserting a paragraph after the one holding it",
        reach: "outside",
        operation: ({ holder }) => ({
          id: "op",
          type: "insertAfterBlock",
          blockId: holder.id,
          text: "Added.",
        }),
      },
    ],
  },
  {
    kind: "block",
    controlText: "Inside text.",
    holderText: "Inside text.",
    bodyXml: (lock) =>
      "<w:p><w:r><w:t>Opening.</w:t></w:r></w:p>" +
      `<w:sdt>${controlProperties(lock)}<w:sdtContent>` +
      "<w:p><w:r><w:t>Inside text.</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
      "<w:p><w:r><w:t>Closing.</w:t></w:r></w:p>",
    cases: [
      {
        name: "replacing text before it",
        reach: "outside",
        operation: ({ before }) => ({
          id: "op",
          type: "replaceInBlock",
          blockId: before.id,
          find: "Opening",
          replace: "Start",
        }),
      },
      {
        name: "deleting the paragraph before it",
        reach: "outside",
        operation: ({ before }) => ({
          id: "op",
          type: "deleteBlock",
          blockId: before.id,
        }),
      },
      {
        name: "replacing its text",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceInBlock",
          blockId: holder.id,
          find: "Inside",
          replace: "Within",
        }),
      },
      {
        name: "formatting its text",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "formatRange",
          range: range(holder, 0, 6),
          formatting: { italic: true },
        }),
      },
      {
        name: "setting its paragraph's properties",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "setBlockParagraphProperties",
          blockId: holder.id,
          properties: { alignment: "center" },
        }),
      },
      {
        name: "splitting its paragraph",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "splitBlock",
          blockId: holder.id,
          offset: 6,
          separator: " ",
        }),
      },
      {
        name: "rewriting its paragraph",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "replaceBlock",
          blockId: holder.id,
          text: "New words.",
        }),
      },
      {
        name: "deleting its paragraph",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "deleteBlock",
          blockId: holder.id,
        }),
      },
      {
        name: "inserting a paragraph after its own",
        reach: "content",
        operation: ({ holder }) => ({
          id: "op",
          type: "insertAfterBlock",
          blockId: holder.id,
          text: "Added.",
        }),
      },
    ],
  },
];

/** The control's `w:sdt` elements in a saved body, with their text. */
const controlsIn = async (saved: ArrayBuffer): Promise<string[]> => {
  const xml =
    (await (await ZipArchive.read(saved)).entry("word/document.xml")?.text()) ??
    "";
  return [...xml.matchAll(/<w:sdt>(.*?)<\/w:sdt>/gu)]
    .map(([, inner]) => inner ?? "")
    .filter((inner) => inner.includes('w:val="Guarded"'))
    .map((inner) =>
      [...inner.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
        .map(([, text]) => text)
        .join(""),
    );
};

const open = async (placement: Placement, lock: Lock) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(placement.bodyXml(lock)),
  );
  const blocks = reviewer.snapshot().blocks;
  const holderIndex = blocks.findIndex(
    ({ text }) => text === placement.holderText,
  );
  const holder = blocks[holderIndex];
  const before = blocks[holderIndex - 1];
  if (!holder || !before) throw new Error("expected the fixture's blocks");
  return { reviewer, holder, before };
};

describe("content-control locks", () => {
  for (const placement of PLACEMENTS)
    for (const lock of LOCKS)
      for (const mode of MODES)
        for (const { name, reach, operation } of placement.cases) {
          const refused = refuses(lock, reach);
          test(`${placement.kind} ${String(lock)}, ${mode}: ${name} is ${refused ? "refused" : "allowed"}`, async () => {
            const { reviewer, holder, before } = await open(placement, lock);
            const original = await controlsIn(await reviewer.toBuffer());
            expect(original).toEqual([placement.controlText]);
            const result = reviewer.applyOperations(
              [operation({ holder, before })],
              { undefinedReferences: "refuse", mode },
            );
            const lockSkips = result.skipped.filter(
              ({ reason }) => reason === "contentControlLock",
            );
            if (refused) {
              expect(result.applied).toEqual([]);
              expect(lockSkips).toHaveLength(1);
              expect(lockSkips[0]?.message).toContain(String(lock));
              expect(await controlsIn(await reviewer.toBuffer())).toEqual(
                original,
              );
              if (mode === "suggested")
                expect(reviewer.exportPendingSuggestions()).toEqual([]);
              return;
            }
            expect(lockSkips).toEqual([]);
            if (mode === "suggested") return;
            // Whatever applied leaves what the lock guards as it was, once
            // every change is accepted.
            reviewer.acceptAll();
            const accepted = await controlsIn(await reviewer.toBuffer());
            if (lock === "contentLocked" || lock === "sdtContentLocked")
              expect(
                accepted.every((text) => text === placement.controlText),
              ).toBe(true);
            if (lock === "sdtLocked" || lock === "sdtContentLocked")
              expect(accepted).toHaveLength(1);
          });
        }
});
