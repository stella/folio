/**
 * Replacing exactly the text a comment covers keeps the comment on the new
 * text, in direct and tracked mode, pending and after accepting, as the
 * default engine does: the replacement's edges stay inside the comment's
 * range markers. Both engines keep the same comment for every placement of
 * the commented words in the paragraph.
 */
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A paragraph of `before`, the commented `word`, then `after`. */
const packageOf = async (before: string, word: string, after: string) => {
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
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="R" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check this.</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="22222222">${before ? run(before) : ""}<w:commentRangeStart w:id="7"/>${run(word)}<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>${after ? run(after) : ""}</w:p>` +
      `<w:p w14:paraId="33333333"/>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** The comments a package reads, pending and once every change is accepted. */
const commentsOf = async (bytes: ArrayBuffer) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
  const pending = reviewer
    .getComments()
    .map(({ id, anchoredText }) => ({ id, anchoredText }));
  reviewer.acceptAll();
  const accepted = reviewer
    .getComments()
    .map(({ id, anchoredText }) => ({ id, anchoredText }));
  return { pending, accepted };
};

const viaDefaultEngine = async (
  source: ArrayBuffer,
  word: string,
  mode: "direct" | "tracked-changes",
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  const result = reviewer.applyOperations(
    [
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: "22222222",
        find: word,
        replace: "Changed",
      },
    ],
    { mode, undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(1);
  return reviewer.toBuffer();
};

const viaCanonicalEngine = async (
  source: ArrayBuffer,
  word: string,
  mode: "direct" | "tracked-changes",
) => {
  const document = await parseDocx(source, {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: {
      version: 1,
      mode,
      operations: [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: "22222222",
          find: word,
          replace: "Changed",
        },
      ],
    },
    undefinedReferences: "refuse",
  });
  expect(result.applied).toHaveLength(1);
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

const PLACEMENTS = [
  ["at the start", "", "Documentation", " of Training"],
  ["in the middle", "Notes on ", "Documentation", " of Training"],
  ["at the end", "Notes on ", "Documentation", ""],
] as const;

describe("replacing exactly the text a comment covers", () => {
  for (const [placement, before, word, after] of PLACEMENTS)
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`${placement}, ${mode}: keeps the comment on the new text in both engines`, async () => {
        const source = await packageOf(before, word, after);
        const expected = await commentsOf(
          await viaDefaultEngine(source, word, mode),
        );
        expect(expected.accepted).toEqual([
          { id: expect.any(Number), anchoredText: "Changed" },
        ]);
        const actual = await commentsOf(
          await viaCanonicalEngine(source, word, mode),
        );
        expect(actual.accepted).toEqual(expected.accepted);
        // While the change is pending the comment covers the struck words,
        // then the new ones, in both engines.
        expect(expected.pending).toEqual([
          {
            id: expected.accepted.at(0)?.id ?? -1,
            anchoredText: mode === "direct" ? "Changed" : `${word}Changed`,
          },
        ]);
        expect(actual.pending).toEqual(expected.pending);
      });
});
