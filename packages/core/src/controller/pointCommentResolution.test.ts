/**
 * A point comment (a bare `w:commentReference`, with no range) lives as long
 * as its reference does. The legacy editor shows it on the text before the
 * reference; deleting that text and accepting the deletion keeps the comment
 * at its reference, as rejecting does, and the save writes it back as a
 * point. Both engines end a review with the same comments.
 */
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { schema } from "../prosemirror/schema";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const deleted = (id: number, text: string) =>
  `<w:del w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const reference = '<w:r><w:commentReference w:id="11"/></w:r>';

/** Where the tracked deletions stand around the point comment's reference. */
const SHAPES = {
  "the word before it": `${run("Keep ")}${deleted(12, "gone")}${reference}${run(" tail")}`,
  "the word after it": `${run("Keep ")}${reference}${deleted(12, "gone")}${run(" tail")}`,
  "the words on both sides": `${run("Keep ")}${deleted(12, "gone")}${reference}${deleted(14, "also")}${run(" tail")}`,
  "everything before it": `${deleted(12, "All of it")}${reference}${run(" tail")}`,
} as const;

const packageOf = async (body: string): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="11" w:author="R" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check this.</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="11111111">${body}</w:p><w:p w14:paraId="22222222"/>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const commentIds = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map(({ id }) => id);

const referencesIn = async (bytes: ArrayBuffer) => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  return {
    references: (xml.match(/<w:commentReference w:id="11"\/>/gu) ?? []).length,
    ranges: (xml.match(/<w:commentRange(?:Start|End) w:id="11"\/>/gu) ?? [])
      .length,
  };
};

/** The canonical session's own accept or reject, saved and read back. */
const canonicallyResolved = async (
  source: ArrayBuffer,
  resolution: "accept" | "reject",
): Promise<ArrayBuffer> => {
  const document = await parseDocx(
    (await prepareCanonicalDocxInput(source)).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  const commit = session
    .prepareResolve(state, { revisionIds: [12, 14], resolution })
    .unwrap();
  expect(publishCanonicalProjection({ state, commit, session }).isOk()).toBe(
    true,
  );
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

describe("a point comment beside tracked deletions", () => {
  for (const [name, body] of Object.entries(SHAPES))
    for (const resolution of ["accept", "reject"] as const)
      test(`next to ${name}: ${resolution}ing keeps the comment at its reference, in both engines`, async () => {
        const source = await packageOf(body);
        const reviewer = await FolioDocxReviewer.fromBuffer(source);
        expect(commentIds(reviewer)).toEqual([11]);
        if (resolution === "accept") reviewer.acceptAll();
        else reviewer.rejectAll();
        expect(commentIds(reviewer)).toEqual([11]);
        const saved = await reviewer.toBuffer();
        // Saved back as the point it was: one reference, no range.
        expect(await referencesIn(saved)).toEqual({ references: 1, ranges: 0 });
        expect(commentIds(await FolioDocxReviewer.fromBuffer(saved))).toEqual([
          11,
        ]);

        const canonical = await canonicallyResolved(source, resolution);
        expect(
          commentIds(await FolioDocxReviewer.fromBuffer(canonical)),
        ).toEqual(commentIds(await FolioDocxReviewer.fromBuffer(saved)));
        expect(await referencesIn(canonical)).toEqual(
          await referencesIn(saved),
        );
      });
});
