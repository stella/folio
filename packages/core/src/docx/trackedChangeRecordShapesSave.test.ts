/**
 * Suggestions over records a tracked change has to hold survive a save:
 * a point comment's reference inside a deleted paragraph's `w:del`, a
 * `w:ins`/`w:del` inside `w:fldSimple`, and a deleted paragraph mark that
 * carries `w:sectPr`. Each reopens as the same pending change and resolves
 * as it did before the save.
 */
import { expect, test } from "bun:test";

import {
  applyDocumentOps,
  compileEditorIntent,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";

const AUTHORED = 'w:author="Earlier" w:date="2026-01-01T00:00:00Z"';
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const NEW_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 101 + index),
};

const open = async (body: string): Promise<Document> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

/** The document as a canonical save writes it, read back. */
const saved = (document: Document): Promise<ArrayBuffer> =>
  repackWithCanonicalStoryRemovals({
    document,
    repack: () => repackDocx(document, { updateModifiedDate: false }),
  });

const reopened = async (document: Document): Promise<Document> =>
  parseDocx(await saved(document), { preloadFonts: false });

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const suggest = (document: Document, intent: EditorIntent) => {
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "suggesting", revision: STAMP, newIds: NEW_IDS },
  });
  if (compiled.isErr()) throw compiled.error;
  return apply(document, compiled.value.ops);
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]).document;

/** The paragraphs' text and the kinds of record they hold, for comparing a document across a save. */
const shape = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          id: block.paraId,
          mark: block.pPrMark?.kind,
          section: block.sectionProperties !== undefined,
          content: JSON.stringify(block.content, (key, value: unknown) =>
            key === "formatting" || key === "date" ? undefined : value,
          ),
        }
      : { id: block.type },
  );

/** The suggestion reopens as itself, and each resolution of it resolves as it did before the save. */
const assertSurvivesSave = async (suggested: {
  document: Document;
  revisions: readonly number[];
}) => {
  const reread = await reopened(suggested.document);
  expect(shape(reread)).toEqual(shape(suggested.document));
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const before = resolve(suggested.document, suggested.revisions, decision);
    const after = resolve(reread, suggested.revisions, decision);
    expect(shape(after)).toEqual(shape(before));
    expect(shape(await reopened(after))).toEqual(shape(before));
  }
  return reread;
};

test("a point comment's reference saves inside the deleted paragraph's w:del", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Noted </w:t></w:r><w:r><w:t>here.</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Kept.</w:t></w:r></w:p>',
  );
  const [first] = document.package.document.content;
  if (first?.type !== "paragraph") throw new Error("Expected a paragraph.");
  first.content.splice(1, 0, { type: "commentReference", id: 7 });
  document.package.document.comments = [
    {
      id: 7,
      author: "Commenter",
      anchorKind: "point",
      content: [
        {
          type: "paragraph",
          paraId: "00000010",
          content: [{ type: "run", content: [{ type: "text", text: "Why?" }] }],
        },
      ],
    },
  ];
  const suggested = suggest(document, {
    type: "deleteParagraphs",
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
  });
  const zip = await ZipArchive.read(await saved(suggested.document));
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  expect(xml).toMatch(
    /<w:del [^>]*>(?:(?!<\/w:del>).)*<w:commentReference w:id="7"\/>(?:(?!<\/w:del>).)*<\/w:del>/su,
  );
  const reread = await assertSurvivesSave(suggested);
  expect(reread.package.document.comments?.map(({ id }) => id)).toEqual([7]);
  const accepted = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  expect(accepted.package.document.comments ?? []).toEqual([]);
  const rejected = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.REJECT,
  );
  expect(rejected.package.document.comments?.map(({ id }) => id)).toEqual([7]);
});

test("a w:ins and a w:del inside w:fldSimple resolve after a save", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r>' +
      `<w:ins w:id="5" ${AUTHORED}><w:r><w:t>2</w:t></w:r></w:ins>` +
      `<w:del w:id="6" ${AUTHORED}><w:r><w:delText>3</w:delText></w:r></w:del>` +
      "</w:fldSimple></w:p>",
  );
  const reread = await assertSurvivesSave({ document, revisions: [5, 6] });
  const fieldText = (resolved: Document) => {
    const [paragraph] = resolved.package.document.content;
    const field =
      paragraph?.type === "paragraph"
        ? paragraph.content.find((item) => item.type === "simpleField")
        : undefined;
    return JSON.stringify(field?.type === "simpleField" ? field.content : [])
      .match(/"text":"[^"]*"/gu)
      ?.join("");
  };
  expect(fieldText(resolve(reread, [5, 6], REVISION_DECISIONS.ACCEPT))).toBe(
    '"text":"1""text":"2"',
  );
  expect(fieldText(resolve(reread, [5, 6], REVISION_DECISIONS.REJECT))).toBe(
    '"text":"1""text":"3"',
  );
});

test("a suggested join at a section break saves its deleted mark beside w:sectPr", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>Ends a section.</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Opens the next.</w:t></w:r></w:p>',
  );
  const suggested = suggest(document, {
    type: "joinParagraphs",
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "00000002",
  });
  const reread = await assertSurvivesSave(suggested);
  const [first] = reread.package.document.content;
  expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe("del");
  expect(first?.type === "paragraph" && first.sectionProperties).toBeTruthy();
  const accepted = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  expect(shape(accepted)).toEqual([
    {
      id: "00000002",
      mark: undefined,
      section: false,
      content: expect.stringContaining("Ends a section.Opens the next."),
    },
  ]);
});
