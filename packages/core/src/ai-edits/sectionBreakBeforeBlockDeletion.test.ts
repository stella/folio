/**
 * Deleting a paragraph that ends a section and stands right before a table
 * or a block content control, with a paragraph before it.
 *
 * A section break lives on the mark of the section's last paragraph
 * (ECMA-376 Part 1 §17.6.18). On canonical, the direct deletion takes the
 * paragraph's words and the mark of the paragraph before it: the paragraph
 * before runs into the deleted paragraph's mark, which keeps its break and
 * its identity, under the paragraph before's properties. Tracked, the
 * deletion records exactly those words and that mark deleted, so accepting
 * it leaves the direct result, rejecting it restores the original, and
 * undoing it restores the original too, saved and reopened.
 *
 * Legacy removes the deleted paragraph's own mark, break and all: its
 * section runs on into the next one, directly and once the tracked deletion
 * is accepted.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { BlockContent, Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const TEXT = {
  lead: "Lead.",
  ends: "Ends one.",
  inside: "Inside.",
  after: "Two.",
} as const;

/** Each paragraph's `w14:paraId`, stated so every reading names it alike. */
const PARA_IDS = {
  [TEXT.lead]: "1A000001",
  [TEXT.ends]: "1A000002",
  [TEXT.inside]: "1A000003",
  [TEXT.after]: "1A000004",
} as const;

/** Portrait ends section one; the body's own section is landscape. */
const SECTION_ONE = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const SECTION_TWO =
  '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>';

const paragraphXml = (
  text: (typeof TEXT)[keyof typeof TEXT],
  properties = "",
): string =>
  `<w:p w14:paraId="${PARA_IDS[text]}">${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const FOLLOWERS = {
  table:
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraphXml(TEXT.inside)}</w:tc></w:tr></w:tbl>`,
  "content control": `<w:sdt><w:sdtPr/><w:sdtContent>${paragraphXml(TEXT.inside)}</w:sdtContent></w:sdt>`,
} as const;
type Follower = keyof typeof FOLLOWERS;

/**
 * "Lead." (centred), "Ends one." (right-aligned, ending section one), the
 * follower, then "Two." in the body's final section.
 */
const fixture = async (follower: Follower): Promise<ArrayBuffer> => {
  const body =
    paragraphXml(TEXT.lead, '<w:jc w:val="center"/>') +
    paragraphXml(TEXT.ends, `<w:jc w:val="right"/>${SECTION_ONE}`) +
    FOLLOWERS[follower] +
    paragraphXml(TEXT.after) +
    SECTION_TWO;
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const open = async (
  bytes: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    author: "Reviewer",
    session,
  });

const reopen = async (
  reviewer: FolioDocxReviewer,
  session: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => await open(await reviewer.toBuffer(), session);

const deleteEnding = (
  reviewer: FolioDocxReviewer,
  mode: FolioAIEditApplyMode,
) => {
  const block = reviewer
    .getContent()
    .find((candidate) => candidate.text === TEXT.ends);
  if (!block) throw new Error(`no block "${TEXT.ends}"`);
  return reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    },
    { undefinedReferences: "refuse" },
  );
};

const body = (reviewer: FolioDocxReviewer): Document["package"]["document"] =>
  reviewer.toDocument().package.document;

/** One top-level block: a paragraph's identity, text, alignment and the section it ends. */
type Projected =
  | {
      kind: "paragraph";
      id: string | undefined;
      text: string;
      alignment: string | undefined;
      ends: string | null;
    }
  | { kind: "block"; type: string };

const textOf = (value: unknown): string => {
  if (typeof value !== "object" || value === null) return "";
  if ("type" in value && value.type === "text" && "text" in value)
    return String(value.text);
  if ("content" in value && Array.isArray(value.content))
    return value.content.map(textOf).join("");
  return "";
};

const orientation = (block: BlockContent): string | null =>
  block.type === "paragraph" && block.sectionProperties !== undefined
    ? (block.sectionProperties.orientation ?? "portrait")
    : null;

const project = (reviewer: FolioDocxReviewer): Projected[] =>
  body(reviewer).content.map((block) =>
    block.type === "paragraph"
      ? {
          kind: "paragraph",
          id: block.paraId,
          text: textOf(block),
          alignment: block.formatting?.alignment,
          ends: orientation(block),
        }
      : { kind: "block", type: block.type },
  );

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

describe.each(Object.keys(FOLLOWERS) as Follower[])(
  "a section-ending paragraph before a %s",
  (follower) => {
    describe.each(SESSIONS)("%s", (session) => {
      test("deleted directly", async () => {
        const source = await fixture(follower);
        const original = project(await open(source, session));
        const [lead, ends, block, after] = original;
        if (lead?.kind !== "paragraph" || ends?.kind !== "paragraph")
          throw new Error("fixture paragraphs missing");

        const reviewer = await open(source, session);
        expect(deleteEnding(reviewer, "direct").applied).toHaveLength(1);

        const expected =
          session === "canonical"
            ? // The paragraph before runs into the deleted paragraph's mark,
              // which keeps its break and identity under the lead's
              // properties.
              [{ ...lead, id: ends.id, ends: ends.ends }, block, after]
            : // Restate legacy: the deleted paragraph's mark goes, break and
              // all, and its section runs on into the next one.
              [lead, block, after];
        expect(project(reviewer)).toEqual(expected);
        expect(project(await reopen(reviewer, session))).toEqual(expected);
      });

      test("deleted tracked: accept, reject and undo", async () => {
        const source = await fixture(follower);
        const original = await open(source, session);
        const originalBlocks = project(original);

        const direct = await open(source, session);
        deleteEnding(direct, "direct");
        const directSaved = await reopen(direct, session);

        const reviewer = await open(source, session);
        const result = deleteEnding(reviewer, "tracked-changes");
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(1);

        // Pending, the section break is still there.
        const pending = await reopen(reviewer, session);
        expect(
          project(pending).filter(
            (block) => block.kind === "paragraph" && block.ends !== null,
          ),
        ).toHaveLength(1);

        const accepting = await reopen(pending, session);
        accepting.acceptAll();
        const accepted = await reopen(accepting, session);
        if (session === "canonical") {
          // Accepting leaves exactly the direct result.
          expect(body(accepted).content).toEqual(body(directSaved).content);
          expect(body(accepted).finalSectionProperties).toEqual(
            body(directSaved).finalSectionProperties,
          );
        } else {
          // Restate legacy: accepting drops the break, as its direct
          // deletion does.
          expect(project(accepted)).toEqual(project(directSaved));
        }

        const rejecting = await reopen(pending, session);
        rejecting.rejectAll();
        expect(body(await reopen(rejecting, session))).toEqual(
          body(await reopen(original, session)),
        );
        expect(project(rejecting)).toEqual(originalBlocks);

        if (result.undoHandle === null) throw new Error("no undo handle");
        expect(reviewer.undoDocumentOperations(result.undoHandle).status).toBe(
          "undone",
        );
        expect(body(await reopen(reviewer, session))).toEqual(
          body(await reopen(original, session)),
        );
      });

      test("suggested, then accepted", async () => {
        const source = await fixture(follower);
        const direct = await open(source, session);
        deleteEnding(direct, "direct");
        const directSaved = await reopen(direct, session);

        const reviewer = await open(source, session);
        const result = deleteEnding(reviewer, "suggested");
        expect(result.skipped).toEqual([]);
        // A suggestion stays out of the package until someone accepts it.
        expect(project(await reopen(reviewer, session))).toEqual(
          project(await open(source, session)),
        );
        for (const { suggestionId } of reviewer.exportPendingSuggestions())
          expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
        reviewer.acceptAll();
        const accepted = await reopen(reviewer, session);
        if (session === "canonical")
          expect(body(accepted).content).toEqual(body(directSaved).content);
        // Restate legacy: accepting drops the break, as its direct deletion
        // does.
        else expect(project(accepted)).toEqual(project(directSaved));
      });
    });
  },
);
