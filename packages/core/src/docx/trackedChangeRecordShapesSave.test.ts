/**
 * Suggestions over records a tracked change has to hold survive a save:
 * a point comment's reference inside a deleted paragraph's `w:del`, a comment
 * anchor in a `w:hyperlink` or `w:sdt` there, a `w:ins`/`w:del` inside
 * `w:fldSimple` (holding boundaries and controls too), a deleted or
 * inserted paragraph mark that carries `w:sectPr` (also with a bookmark
 * standing between it and the next paragraph), a `w:sectPrChange` on the
 * body's final `w:sectPr`, a
 * `w:rPrChange` on a paragraph mark's `w:pPr/w:rPr`, an inserted mark a
 * direct paragraph took over, and a pasted break's `w:ins` holding a copied
 * `w:del` or `w:moveFrom` beside it. Each reopens as the same pending change and resolves as
 * it did before the save.
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

import type {
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";

const AUTHORED = 'w:author="Earlier" w:date="2026-01-01T00:00:00Z"';
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const NEW_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 101 + index),
};

const open = async (
  body: string,
  finalSection = "<w:sectPr/>",
): Promise<Document> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${finalSection}</w:body></w:document>`,
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

test.each([
  [
    "w:hyperlink",
    '<w:hyperlink w:anchor="target"><w:r><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r></w:hyperlink>',
  ],
  [
    "w:sdt",
    '<w:sdt><w:sdtPr><w:id w:val="55"/></w:sdtPr><w:sdtContent><w:r><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r></w:sdtContent></w:sdt>',
  ],
])(
  "a comment range start inside %s stands outside the deleted paragraph's w:del after a save",
  async (_container, held) => {
    const document = await open(
      `<w:p w14:paraId="00000001"><w:r><w:t>a</w:t></w:r>${held}<w:r><w:t>d</w:t></w:r></w:p>` +
        '<w:p w14:paraId="00000002"><w:r><w:t>e</w:t></w:r></w:p>',
    );
    const [first, second] = document.package.document.content;
    if (first?.type !== "paragraph" || second?.type !== "paragraph")
      throw new Error("Expected two paragraphs.");
    const container = first.content[1];
    if (container?.type === "hyperlink")
      container.children.splice(1, 0, { type: "commentRangeStart", id: 7 });
    else if (container?.type === "inlineSdt")
      container.content.splice(1, 0, { type: "commentRangeStart", id: 7 });
    else throw new Error("Expected a hyperlink or a content control.");
    second.content.push(
      { type: "commentRangeEnd", id: 7 },
      { type: "commentReference", id: 7 },
    );
    document.package.document.comments = [
      {
        id: 7,
        author: "Commenter",
        content: [
          {
            type: "paragraph",
            paraId: "00000010",
            content: [
              { type: "run", content: [{ type: "text", text: "Why?" }] },
            ],
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
    expect(xml).not.toMatch(
      /<w:del [^>]*>(?:(?!<\/w:del>).)*<w:commentRangeStart w:id="7"\/>/su,
    );
    expect(xml).toContain('<w:commentRangeStart w:id="7"/>');
    const reread = await assertSurvivesSave(suggested);
    expect(reread.package.document.comments?.map(({ id }) => id)).toEqual([7]);
  },
);

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

test.each([REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])(
  "a change inside w:fldSimple holding a bookmark and a control resolves in place across a save (%s)",
  async (decision) => {
    const document = await open(
      '<w:p w14:paraId="00000001"><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r>' +
        `<w:ins w:id="5" ${AUTHORED}><w:bookmarkStart w:id="3" w:name="_Ref3"/><w:r><w:t>2</w:t></w:r>` +
        '<w:sdt><w:sdtContent><w:r><w:t>9</w:t></w:r></w:sdtContent></w:sdt><w:bookmarkEnd w:id="3"/></w:ins>' +
        "</w:fldSimple></w:p>",
    );
    const written = (from: Document): string => {
      const [paragraph] = from.package.document.content;
      if (paragraph?.type !== "paragraph")
        throw new TypeError("The document opens with the paragraph.");
      return serializeParagraph(paragraph);
    };
    const resolved = resolve(document, [5], decision);
    expect(written(resolve(await reopened(document), [5], decision))).toBe(
      written(resolved),
    );
    expect(written(await reopened(resolved))).toBe(written(resolved));
    // The change's content stays in the field, or goes leaving the
    // bookmark's boundaries where they stood.
    expect(written(resolved)).toContain(
      decision === REVISION_DECISIONS.ACCEPT
        ? '<w:r><w:t>1</w:t></w:r><w:bookmarkStart w:id="3" w:name="_Ref3"/><w:r><w:t>2</w:t></w:r><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>9</w:t></w:r></w:sdtContent></w:sdt><w:bookmarkEnd w:id="3"/></w:fldSimple>'
        : '<w:r><w:t>1</w:t></w:r><w:bookmarkStart w:id="3" w:name="_Ref3"/><w:bookmarkEnd w:id="3"/></w:fldSimple>',
    );
  },
);

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

test("a final section revision resolves after a save", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t>Body.</w:t></w:r></w:p>',
    '<w:sectPr><w:pgSz w:w="14400" w:h="15840"/>' +
      `<w:sectPrChange w:id="9" ${AUTHORED}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>`,
  );
  const reread = await reopened(document);
  const width = (resolved: Document) =>
    resolved.package.document.finalSectionProperties?.pageWidth;
  for (const source of [document, reread]) {
    const accepted = resolve(source, [9], REVISION_DECISIONS.ACCEPT);
    const rejected = resolve(source, [9], REVISION_DECISIONS.REJECT);
    expect(width(accepted)).toBe(14_400);
    expect(width(rejected)).toBe(12_240);
    for (const resolved of [accepted, rejected]) {
      expect(
        resolved.package.document.finalSectionProperties?.propertyChanges,
      ).toBeUndefined();
      expect(width(await reopened(resolved))).toBe(width(resolved));
    }
  }
});

test("a suggested paragraph ending a section saves its break beside its inserted mark", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t>One.</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Two.</w:t></w:r></w:p>',
  );
  const suggested = apply(document, [
    {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      blocks: [
        {
          type: "paragraph",
          paraId: "0000000A",
          sectionProperties: { pageWidth: 11_906, pageHeight: 16_838 },
          content: [
            { type: "run", content: [{ type: "text", text: "Ends." }] },
          ],
        },
      ],
      revision: STAMP,
      newIds: NEW_IDS,
    },
  ]);
  const reread = await assertSurvivesSave(suggested);
  const inserted = reread.package.document.content.find(
    (block) => block.type === "paragraph" && block.paraId === "0000000A",
  );
  expect(inserted?.type === "paragraph" && inserted.pPrMark?.kind).toBe("ins");
  expect(
    inserted?.type === "paragraph" && inserted.sectionProperties?.pageWidth,
  ).toBe(11_906);
  const rejected = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.REJECT,
  );
  expect(shape(rejected)).toEqual(shape(document));
});

test.each(["outside", "inside a w:hyperlink"] as const)(
  "a deletion a comment range crosses %s saves the standing endpoint outside its w:del",
  async (placement) => {
    // Alternate run formatting keeps each letter its own run across a save.
    const letter = (text: string, bold: boolean) =>
      `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t>${text}</w:t></w:r>`;
    const held = letter("b", true) + letter("c", false);
    const document = await open(
      `<w:p w14:paraId="00000001">${letter("a", false)}` +
        (placement === "outside"
          ? held
          : `<w:hyperlink w:anchor="Clause">${held}</w:hyperlink>`) +
        `${letter("d", true)}</w:p>`,
    );
    const [first] = document.package.document.content;
    if (first?.type !== "paragraph")
      throw new TypeError("Expected a paragraph.");
    // The range opens between "b" and "c", inside the deleted "bc", and
    // closes after "d", outside it.
    const start = { type: "commentRangeStart", id: 7 } as const;
    const container = first.content[1];
    if (placement === "outside") first.content.splice(2, 0, start);
    else if (container?.type === "hyperlink")
      container.children.splice(1, 0, start);
    else throw new TypeError("Expected a hyperlink.");
    first.content.push(
      { type: "commentRangeEnd", id: 7 },
      { type: "commentReference", id: 7 },
    );
    document.package.document.comments = [
      { id: 7, author: "Commenter", anchorKind: "range", content: [] },
    ];
    const suggested = apply(document, [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
        to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 3 },
        revision: STAMP,
        newIds: NEW_IDS,
      },
    ]);
    const zip = await ZipArchive.read(await saved(suggested.document));
    const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
    expect(xml).toContain("<w:del ");
    expect(xml).not.toMatch(
      /<w:del [^>]*>(?:(?!<\/w:del>).)*<w:commentRangeStart w:id="7"\/>/su,
    );
    const reread = await assertSurvivesSave(suggested);
    expect(reread.package.document.comments?.map(({ id }) => id)).toEqual([7]);
  },
);

test("a suggested join through hyperlinks saves its merge depth on the deleted mark", async () => {
  const link = (text: string) =>
    `<w:hyperlink w:anchor="Clause"><w:r><w:t>${text}</w:t></w:r></w:hyperlink>`;
  const document = await open(
    `<w:p w14:paraId="00000001">${link("a")}</w:p>` +
      `<w:p w14:paraId="00000002">${link("b")}</w:p>`,
  );
  const suggested = apply(document, [
    {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      depth: 2,
      revision: STAMP,
    },
  ]);
  const reread = await assertSurvivesSave(suggested);
  const [first] = reread.package.document.content;
  expect(first?.type === "paragraph" && first.pPrMark?.resolutionJoin).toBe(2);
  const accepted = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  const [joined] = accepted.package.document.content;
  if (joined?.type !== "paragraph")
    throw new TypeError("Expected a paragraph.");
  expect(joined.content.map(({ type }) => type)).toEqual(["hyperlink"]);
});

test("a suggested paragraph keeps a pasted control's own insertion after a save", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t>Here.</w:t></w:r></w:p>',
  );
  const pasted: Paragraph = {
    type: "paragraph",
    paraId: "0000000B",
    content: [
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 42 },
        content: [
          {
            type: "insertion",
            info: { id: 91, author: "Earlier", date: "2026-01-01T00:00:00Z" },
            content: [
              { type: "run", content: [{ type: "text", text: "pasted" }] },
            ],
          },
        ],
      },
    ],
  };
  const suggested = suggest(document, {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000001" },
    paragraphs: [pasted],
  });
  expect(suggested.revisions).not.toContain(91);
  const reread = await assertSurvivesSave(suggested);
  const rejected = resolve(
    reread,
    suggested.revisions,
    REVISION_DECISIONS.REJECT,
  );
  expect(shape(rejected)).toEqual(shape(document));
});

test("a property change on a field's instruction run resolves after a save", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="8" ${AUTHORED}><w:rPr/></w:rPrChange></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  );
  const reread = await assertSurvivesSave({ document, revisions: [8] });
  const instruction = (resolved: Document) => {
    const [paragraph] = resolved.package.document.content;
    const field =
      paragraph?.type === "paragraph"
        ? paragraph.content.find((item) => item.type === "complexField")
        : undefined;
    if (field?.type !== "complexField")
      throw new TypeError("Expected a complex field.");
    return field.fieldCode.map(({ formatting, propertyChanges }) => ({
      bold: formatting?.bold === true,
      pending: (propertyChanges ?? []).length,
    }));
  };
  expect(instruction(reread)).toEqual([{ bold: true, pending: 1 }]);
  expect(instruction(resolve(reread, [8], REVISION_DECISIONS.ACCEPT))).toEqual([
    { bold: true, pending: 0 },
  ]);
  expect(instruction(resolve(reread, [8], REVISION_DECISIONS.REJECT))).toEqual([
    { bold: false, pending: 0 },
  ]);
});

const MARK_PATCHES: [string, ParagraphFormatting][] = [
  ["the mark's run properties", { runProperties: { italic: true } }],
  [
    "paragraph and mark properties",
    { alignment: "center", runProperties: { italic: true } },
  ],
];

test.each(MARK_PATCHES)(
  "a suggested change to %s saves its mark record under w:pPr/w:rPr and resolves after a save",
  async (_name, patch) => {
    const document = await open(
      '<w:p w14:paraId="00000001"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
        '<w:p w14:paraId="00000002"><w:r><w:t>Body.</w:t></w:r></w:p>',
    );
    const suggested = suggest(document, {
      type: "formatParagraph",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      patch,
    });
    const zip = await ZipArchive.read(await saved(suggested.document));
    const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
    expect(xml).toMatch(
      /<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr><w:i\/><w:rPrChange [^>]*><w:rPr><w:b\/><\/w:rPr><\/w:rPrChange><\/w:rPr>/su,
    );
    const reread = await assertSurvivesSave(suggested);
    const mark = (resolved: Document) => {
      const [paragraph] = resolved.package.document.content;
      if (paragraph?.type !== "paragraph")
        throw new TypeError("Expected a paragraph.");
      return {
        alignment: paragraph.formatting?.alignment ?? "none",
        runProperties: paragraph.formatting?.runProperties ?? {},
        previous: paragraph.markPropertyChange?.previousFormatting ?? {},
        paragraphChanges: paragraph.propertyChanges?.length ?? 0,
      };
    };
    expect(mark(reread)).toEqual(mark(suggested.document));
    expect(mark(reread).previous).toEqual({ runProperties: { bold: true } });
    for (const decision of Object.values(REVISION_DECISIONS))
      expect(mark(resolve(reread, suggested.revisions, decision))).toEqual(
        mark(resolve(suggested.document, suggested.revisions, decision)),
      );
    expect(
      mark(resolve(reread, suggested.revisions, REVISION_DECISIONS.REJECT)),
    ).toEqual(mark(document));
  },
);

test("a direct paragraph after a pending inserted mark saves the break it took over", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t>Before.</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Anchor.</w:t></w:r></w:p>',
  );
  const marked = apply(document, [
    {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000002" },
      blocks: [
        {
          type: "paragraph",
          paraId: "00000010",
          content: [{ type: "run", content: [{ type: "text", text: "New." }] }],
        },
      ],
      revision: STAMP,
      newIds: NEW_IDS,
    },
  ]);
  const inserted = apply(marked.document, [
    {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000002" },
      blocks: [
        {
          type: "paragraph",
          paraId: "00000020",
          content: [
            { type: "run", content: [{ type: "text", text: "Direct." }] },
          ],
        },
      ],
    },
  ]);
  const reread = await assertSurvivesSave({
    document: inserted.document,
    revisions: marked.revisions,
  });
  expect(
    shape(reread).map((entry) => ("mark" in entry ? entry.mark : "none")),
  ).toEqual([undefined, undefined, "ins", undefined]);
  expect(
    shape(resolve(reread, marked.revisions, REVISION_DECISIONS.REJECT)).map(
      ({ id }) => id,
    ),
  ).toEqual(["00000001", "00000002", "00000020"]);
});

test.each(["del", "moveFrom"] as const)(
  "a suggested paste keeps a copied %s of a pasted break beside its insertion after a save",
  async (kind) => {
    const document = await open(
      '<w:p w14:paraId="00000001"><w:r><w:t>Here.</w:t></w:r></w:p>',
    );
    const suggested = suggest(document, {
      type: "replaceFragment",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
      openStart: 0,
      openEnd: 0,
      paragraphs: [
        {
          type: "paragraph",
          paraId: "0000000D",
          pPrMark: {
            kind,
            info: { id: 93, author: "Earlier", date: "2026-01-01T00:00:00Z" },
          },
          content: [
            { type: "run", content: [{ type: "text", text: "Copied" }] },
          ],
        },
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "tail" }] }],
        },
      ],
    });
    const reread = await assertSurvivesSave(suggested);
    const deletions = (from: Document) =>
      from.package.document.content.flatMap((block) =>
        block.type === "paragraph" && block.pPrMark?.deletion !== undefined
          ? [
              [
                block.pPrMark.deletion.kind ?? "del",
                block.pPrMark.deletion.info.author,
              ],
            ]
          : [],
      );
    expect(deletions(reread)).toEqual([[kind, "Earlier"]]);
    const accepted = resolve(
      reread,
      suggested.revisions,
      REVISION_DECISIONS.ACCEPT,
    );
    expect(
      accepted.package.document.content.flatMap((block) =>
        block.type === "paragraph" && block.pPrMark !== undefined
          ? [[block.pPrMark.kind, block.pPrMark.info.author]]
          : [],
      ),
    ).toEqual([[kind, "Earlier"]]);
  },
);

test("a suggested paste of a review carrier saves the plain paragraph it shows", async () => {
  const document = await open(
    '<w:p w14:paraId="00000001"><w:r><w:t>Here.</w:t></w:r></w:p>',
  );
  const suggested = suggest(document, {
    type: "replaceFragment",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    openStart: 0,
    openEnd: 0,
    paragraphs: [
      {
        type: "paragraph",
        paraId: "0000000E",
        reviewCarrier: "terminal-table",
        content: [
          { type: "run", content: [{ type: "text", text: "Carried" }] },
        ],
      },
      {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "tail" }] }],
      },
    ],
  });
  const reread = await assertSurvivesSave(suggested);
  expect(
    reread.package.document.content.some(
      (block) =>
        block.type === "paragraph" && block.reviewCarrier !== undefined,
    ),
  ).toBe(false);
});

test("a deleted section-ending mark before a standing bookmark resolves after a save", async () => {
  const document = await open(
    `<w:p w14:paraId="00000001"><w:pPr><w:rPr><w:del w:id="7" ${AUTHORED}/></w:rPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>Ends a section.</w:t></w:r></w:p>` +
      '<w:bookmarkStart w:id="2" w:name="Later"/>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Opens the next.</w:t></w:r></w:p>' +
      '<w:bookmarkEnd w:id="2"/>',
  );
  const reread = await assertSurvivesSave({ document, revisions: [7] });
  const accepted = resolve(reread, [7], REVISION_DECISIONS.ACCEPT);
  const [joined] = accepted.package.document.content;
  if (joined?.type !== "paragraph")
    throw new TypeError("The joined block is a paragraph.");
  expect(joined.sectionProperties).toBeUndefined();
  expect(joined.content.map(({ type }) => type)).toEqual([
    "run",
    "bookmarkStart",
    "run",
  ]);
  expect(shape(await reopened(accepted))).toEqual(shape(accepted));
});
