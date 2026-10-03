/** Package-shaped cases for the same inverse/locality laws across every operation kind. */
import { panic } from "better-result";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../packages/docx-core/src/model/document";
import { applyDocumentOp } from "../../packages/docx-core/src/ops/apply";
import { storyBody, storyParagraphs } from "../../packages/docx-core/src/ops/blocks";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";
import { paragraphIdsIn } from "../../packages/docx-core/src/ops/ids";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
} from "../../packages/docx-core/src/ops/types";
import {
  GENERATED_OP_KINDS,
  opForStory,
  opSeedArbitrary,
  type OpSeed,
} from "../../packages/docx-core/src/ops/__tests__/documentArbitraries";

export const GENERATED_PACKAGE_STORIES = [
  "main",
  { kind: "header", rId: "rIdHeader" },
  { kind: "footer", rId: "rIdFooter" },
  { kind: "footnote", id: 2 },
  { kind: "endnote", id: 3 },
] as const satisfies readonly OpStory[];

const paragraph = (id: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId: id,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** Every draw carries all editable stories and balanced, package-wide identities. */
export const packageDocumentArbitrary = opSeedArbitrary.map((seed): Document => {
  let nextParagraph = 1;
  const named = (text: string): Paragraph =>
    paragraph((nextParagraph++).toString(16).toUpperCase().padStart(8, "0"), text);
  const blocks = (index: number): BlockContent[] => {
    const first = named("xy" + seed.text);
    first.formatting = { alignment: seed.inherit ? "center" : "start", keepNext: true };
    const second = named("follower " + seed.text);
    const reviewed = named(seed.text);
    // Derived from existing seed fields: adding review coverage does not change generation or shrinking.
    if (seed.inherit) {
      reviewed.formatting = {
        alignment: "end",
        preserved: { children: [{ index: 3, xml: "<w:suppressOverlap/>" }] },
      };
      reviewed.propertyChanges = [
        {
          type: "paragraphPropertyChange",
          info: { id: 300 + index, author: "B", date: "2026-01-02T03:04:05Z" },
          previousFormatting: { alignment: "start", keepNext: true },
        },
      ];
    }
    reviewed.content = [
      {
        type: "insertion",
        info: { id: 100 + index, author: "A", date: "2026-01-02T03:04:05Z" },
        content: [
          {
            type: "run",
            formatting: seed.formatting,
            ...(seed.first % 2 === 0
              ? {
                  propertyChanges: [
                    {
                      type: "runPropertyChange" as const,
                      info: { id: 400 + index, author: "B", date: "2026-01-02T03:04:05Z" },
                      previousFormatting: {
                        italic: true,
                        preserved: { children: [{ index: 38, xml: "<w:webHidden/>" }] },
                      },
                    },
                  ],
                }
              : {}),
            content: [{ type: "text", text: seed.text }],
          },
        ],
      },
    ];
    let controlId = 200 + index;
    if (seed.first % 3 === 0) controlId = -(200 + index);
    if (seed.first % 3 === 1) controlId = 0x7f10cd00 + index;
    const control: BlockContent = {
      type: "blockSdt",
      properties: {
        sdtType: "richText",
        // CT_SdtPr uses a signed decimal id; parsed packages include negative and large ids.
        id: controlId,
        tag: "clause",
      },
      content: [named(seed.text)],
    };
    const table: Table = {
      type: "table",
      columnWidths: [2400, 2400],
      rows: [0, 1].map(() => ({
        type: "tableRow",
        cells: [0, 1].map(() => ({ type: "tableCell", content: [named(seed.text)] })),
      })),
    };
    return [first, second, reviewed, control, table, named("tail")];
  };
  const content = blocks(0);
  const mainFirst = content.at(0);
  if (mainFirst?.type !== "paragraph") return panic("Package fixture needs its leading paragraph.");
  mainFirst.content.push(
    { type: "commentRangeStart", id: 7 },
    { type: "run", content: [{ type: "text", text: "comment" }] },
    { type: "commentRangeEnd", id: 7 },
    { type: "commentReference", id: 7 },
  );
  const last = content.at(-1);
  if (last?.type !== "paragraph") return panic("Package fixture needs its final paragraph.");
  last.content.push({
    type: "run",
    content: [
      { type: "footnoteRef", id: 2 },
      { type: "endnoteRef", id: 3 },
    ],
  });
  const header = { type: "header", hdrFtrType: "default", content: blocks(1) } as const;
  const footer = { type: "footer", hdrFtrType: "default", content: blocks(2) } as const;
  const footnote = { type: "footnote", id: 2, content: blocks(3) } as const;
  const endnote = { type: "endnote", id: 3, content: blocks(4) } as const;
  const comment = { id: 7, author: "A", content: [named("comment body")] };
  const properties = {
    pageWidth: 12240,
    pageHeight: 15840,
    headerReferences: [{ type: "default", rId: "rIdHeader" }] as const,
    footerReferences: [{ type: "default", rId: "rIdFooter" }] as const,
  };
  return normalizeForOps({
    package: {
      document: {
        content,
        comments: [comment],
        finalSectionProperties: {
          ...properties,
          headerReferences: [...properties.headerReferences],
          footerReferences: [...properties.footerReferences],
        },
        sections: [
          {
            properties: {
              ...properties,
              headerReferences: [...properties.headerReferences],
              footerReferences: [...properties.footerReferences],
            },
            content,
            headers: new Map([["default", header] as const]),
            footers: new Map([["default", footer] as const]),
          },
        ],
      },
      headers: new Map([["rIdHeader", header]]),
      footers: new Map([["rIdFooter", footer]]),
      footnotes: [footnote],
      endnotes: [endnote],
      settings: { defaultTabStop: 720 },
      properties: {
        title: "operation fixture",
        created: new Date("2026-01-02T03:04:05Z"),
        modified: new Date("2026-01-02T03:04:05Z"),
      },
    },
  });
});

type CaseArgs = { document: Document; seed: OpSeed; story: OpStory };
type GeneratedCase = { document: Document; op: DocumentOp };

const targetFor = ({ document, story }: CaseArgs) => {
  const target = storyParagraphs(storyBody(document, story)).at(0)?.paragraph;
  if (!target?.paraId)
    return panic("A package fixture needs an identified paragraph in every story.");
  return target;
};

const freshParagraph = ({ document, seed }: CaseArgs, delta = 0): Paragraph => {
  const used = new Set(paragraphIdsIn(document.package));
  let value = seed.fresh + delta;
  let id = value.toString(16).toUpperCase().padStart(8, "0");
  while (used.has(id)) id = (++value).toString(16).toUpperCase().padStart(8, "0");
  return paragraph(id, seed.text);
};

const tableFor = ({ document, story }: CaseArgs) => {
  const table = storyBody(document, story).content.find((block) => block.type === "table");
  if (!table) return panic("A package fixture needs a table in every story.");
  const blockId = table.rows.at(0)?.cells.at(0)?.content.at(0);
  if (blockId?.type !== "paragraph" || !blockId.paraId)
    return panic("Table fixture needs an identified cell paragraph.");
  return { table, blockId: blockId.paraId };
};

const direct =
  (kind: DocumentOp["type"]) =>
  (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: opForStory({
      ...args,
      seed: {
        ...args.seed,
        kind: GENERATED_OP_KINDS.indexOf(kind),
        block: 0,
        third: 0,
        inherit: false,
        content: [{ type: "run", content: [{ type: "text", text: args.seed.text }] }],
      },
    }),
  });

const inverseCase = (
  args: CaseArgs,
  forward: DocumentOp,
  kind: DocumentOp["type"],
): GeneratedCase => {
  const applied = applyDocumentOp(args.document, forward).unwrap();
  const op = applied.inverse.find((inverse) => inverse.type === kind);
  if (!op) return panic("The prerequisite edit must produce the requested inverse primitive.");
  return { document: applied.document, op };
};

/** Total by the actual op union, rather than a separately maintained kind list. */
export const PACKAGE_OP_CASES = {
  insertText: direct("insertText"),
  insertContent: direct("insertContent"),
  deleteRange: (args: CaseArgs): GeneratedCase => {
    const blockId = targetFor(args).paraId ?? panic("Missing paragraph identity.");
    return {
      document: args.document,
      op: {
        type: "deleteRange",
        from: { story: args.story, blockId, offset: 0 },
        to: { story: args.story, blockId, offset: 1 },
      },
    };
  },
  splitInline: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "splitInline",
      at: { story: args.story, blockId: targetFor(args).paraId ?? "", offset: 1 },
      depth: 1,
    },
  }),
  joinInline: (args: CaseArgs): GeneratedCase =>
    inverseCase(
      args,
      {
        type: "splitInline",
        at: { story: args.story, blockId: targetFor(args).paraId ?? "", offset: 1 },
        depth: 1,
      },
      "joinInline",
    ),
  setRunProps: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "setRunProps",
      from: { story: args.story, blockId: targetFor(args).paraId ?? "", offset: 0 },
      to: { story: args.story, blockId: targetFor(args).paraId ?? "", offset: 1 },
      patch: { bold: true },
    },
  }),
  setParagraphProps: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "setParagraphProps",
      story: args.story,
      blockId: targetFor(args).paraId ?? "",
      patch: { keepNext: false },
    },
  }),
  splitBlock: direct("splitBlock"),
  joinBlocks: (args: CaseArgs): GeneratedCase => {
    const paragraphs = storyParagraphs(storyBody(args.document, args.story));
    const next = paragraphs.at(1)?.paragraph;
    if (!next?.paraId) return panic("A package fixture needs adjacent leading paragraphs.");
    return {
      document: args.document,
      op: {
        type: "joinBlocks",
        story: args.story,
        blockId: targetFor(args).paraId ?? "",
        nextBlockId: next.paraId,
        depth: 0,
      },
    };
  },
  replaceBlocks: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "replaceBlocks",
      story: args.story,
      expected: [targetFor(args)],
      blocks: [freshParagraph(args)],
    },
  }),
  insertBlocks: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "insertBlocks",
      story: args.story,
      at: { type: "before", blockId: targetFor(args).paraId ?? "" },
      blocks: [freshParagraph(args)],
    },
  }),
  deleteBlocks: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: { type: "deleteBlocks", story: args.story, blockIds: [targetFor(args).paraId ?? ""] },
  }),
  setParagraphReview: (args: CaseArgs): GeneratedCase => {
    const target = targetFor(args);
    return {
      document: args.document,
      op: {
        type: "setParagraphReview",
        story: args.story,
        blockId: target.paraId ?? "",
        expected: target.formatting ? { formatting: target.formatting } : {},
        review: { formatting: { alignment: "end" } },
      },
    };
  },
  replaceInline: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "replaceInline",
      story: args.story,
      blockId: targetFor(args).paraId ?? "",
      expected: targetFor(args).content,
      content: [{ type: "run", content: [{ type: "text", text: args.seed.text }] }],
    },
  }),
  resolveRevision: (args: CaseArgs): GeneratedCase => {
    const index = GENERATED_PACKAGE_STORIES.findIndex(
      (story) => JSON.stringify(story) === JSON.stringify(args.story),
    );
    if (index < 0) return panic("Unknown package fixture story.");
    return {
      document: args.document,
      op: {
        type: "resolveRevision",
        story: args.story,
        revisionIds: [100 + index],
        decision: args.seed.inherit ? "accept" : "reject",
      },
    };
  },
  insertTable: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "insertTable",
      story: args.story,
      at: { type: "before", blockId: targetFor(args).paraId ?? "" },
      table: {
        type: "table",
        columnWidths: [2400],
        rows: [
          { type: "tableRow", cells: [{ type: "tableCell", content: [freshParagraph(args)] }] },
        ],
      },
    },
  }),
  deleteTable: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: { type: "deleteTable", story: args.story, blockId: tableFor(args).blockId },
  }),
  setContainerBlocks: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "setContainerBlocks",
      story: args.story,
      blockId: targetFor(args).paraId ?? "",
      expected: storyBody(args.document, args.story).content,
      blocks: [...storyBody(args.document, args.story).content, freshParagraph(args)],
    },
  }),
  insertRow: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "insertRow",
      story: args.story,
      blockId: tableFor(args).blockId,
      at: args.seed.first % 3,
      row: {
        type: "tableRow",
        cells: [
          { type: "tableCell", content: [freshParagraph(args)] },
          { type: "tableCell", content: [freshParagraph(args, 1)] },
        ],
      },
    },
  }),
  deleteRow: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: { type: "deleteRow", story: args.story, blockId: tableFor(args).blockId },
  }),
  setTableRows: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "setTableRows",
      story: args.story,
      blockId: tableFor(args).blockId,
      expected: tableFor(args).table.rows,
      rows: tableFor(args).table.rows.slice(0, 1),
    },
  }),
  createHeaderFooter: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "createHeaderFooter",
      sectionIndex: 0,
      story: { kind: args.seed.inherit ? "header" : "footer", rId: "rIdGenerated" },
      referenceType: "first",
      content: [freshParagraph(args)],
    },
  }),
  removeHeaderFooter: (args: CaseArgs): GeneratedCase => {
    if (args.seed.first % 2 !== 0)
      return {
        document: args.document,
        op: {
          type: "removeHeaderFooter",
          sectionIndex: 0,
          story: { kind: "header", rId: "rIdHeader" },
          referenceType: "default",
        },
      };
    const created = applyDocumentOp(args.document, {
      type: "createHeaderFooter",
      sectionIndex: 0,
      story: { kind: "footer", rId: "rIdGenerated" },
      referenceType: "first",
      content: [freshParagraph(args)],
    }).unwrap().document;
    return {
      document: created,
      op: {
        type: "removeHeaderFooter",
        sectionIndex: 0,
        story: { kind: "footer", rId: "rIdGenerated" },
        referenceType: "first",
      },
    };
  },
  addNote: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "addNote",
      at: { story: "main", blockId: targetFor({ ...args, story: "main" }).paraId ?? "", offset: 0 },
      note: {
        type: args.seed.inherit ? "footnote" : "endnote",
        id: 19,
        content: [freshParagraph(args)],
      },
    },
  }),
  removeNote: (args: CaseArgs): GeneratedCase => {
    const main = storyParagraphs(storyBody(args.document, "main"));
    const target = main.at(-1)?.paragraph;
    if (!target?.paraId) return panic("Package fixture needs its note-reference paragraph.");
    return {
      document: args.document,
      op: {
        type: "removeNote",
        at: { story: "main", blockId: target.paraId, offset: 4 },
        story: { kind: "footnote", id: 2 },
      },
    };
  },
  setSectionProps: (args: CaseArgs): GeneratedCase => ({
    document: args.document,
    op: {
      type: "setSectionProps",
      sectionIndex: 0,
      patch: { pageWidth: 13000 + (args.seed.first % 100) },
    },
  }),
  restoreStoryParts: (args: CaseArgs): GeneratedCase =>
    inverseCase(
      args,
      { type: "setSectionProps", sectionIndex: 0, patch: { pageWidth: 13000 } },
      "restoreStoryParts",
    ),
} satisfies Record<DocumentOp["type"], (args: CaseArgs) => GeneratedCase>;

export const GENERATED_PACKAGE_OP_KINDS = Object.values(DOCUMENT_OP_TYPES);

type GeneratedCaseArgs = CaseArgs & { kind: DocumentOp["type"] };
export const generatedCaseFor = ({ kind, ...args }: GeneratedCaseArgs): GeneratedCase =>
  PACKAGE_OP_CASES[kind](args);
