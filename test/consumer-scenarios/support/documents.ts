/**
 * Synthetic documents built the way an integrator builds them: Markdown
 * through `fromMarkdown`, a few model edits through the published builders,
 * `createDocx`, then `ensureParaIds`. Each fixture returns package bytes; the
 * edits that need a reviewer (comments, tracked changes) go through the public
 * operation contract and a save, so a fixture is itself a first round trip.
 */

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  endnote,
  ensureParaIds,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
  paragraph,
  run,
  table,
} from "@stll/folio-core/server";

export type FolioDocument = ReturnType<typeof fromMarkdown>;
type Numbering = NonNullable<FolioDocument["package"]["numbering"]>;
type NumberingLevel = Numbering["abstractNums"][number]["levels"][number];
type Block = FolioDocument["package"]["document"]["content"][number];
type Paragraph = Extract<Block, { type: "paragraph" }>;
type Reviewer = Awaited<ReturnType<typeof FolioDocxReviewer.fromBuffer>>;
type OperationBatch = Parameters<Reviewer["applyDocumentOperations"]>[0];

export const AUTHOR = "Consumer Scenario";

export const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/** `createDocx` + `ensureParaIds`, the recipe the published README gives. */
export const packDocument = async (document: FolioDocument): Promise<Uint8Array> =>
  (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;

export const openReviewer = (bytes: Uint8Array, author = AUTHOR): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes), { author });

export const decimalLevel = (ilvl: number, lvlText: string): NumberingLevel => ({
  ilvl,
  start: 1,
  numFmt: "decimal",
  lvlText,
  suffix: "space",
  pPr: { indentLeft: 0, indentFirstLine: 0 },
});

const findStyle = (document: FolioDocument, styleId: string) => {
  const style = document.package.styles?.styles.find((candidate) => candidate.styleId === styleId);
  if (!style) {
    throw new Error(`fixture style ${styleId} is missing`);
  }
  return style;
};

const findParagraph = (document: FolioDocument, text: string): Paragraph => {
  const found = document.package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" &&
      block.content.some(
        (item) =>
          item.type === "run" &&
          item.content.some((content) => content.type === "text" && content.text === text),
      ),
  );
  if (!found) {
    throw new Error(`fixture paragraph "${text}" is missing`);
  }
  return found;
};

const blockIdOf = (reviewer: Reviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`fixture block "${text}" is missing`);
  }
  return block.id;
};

const applyOrThrow = (reviewer: Reviewer, batch: Omit<OperationBatch, "version">): void => {
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    ...batch,
  } as OperationBatch);
  if (result.skipped.length > 0 || result.issues.length > 0) {
    throw new Error(`fixture operations were refused: ${JSON.stringify(result.issues)}`);
  }
};

const PLAIN_MARKDOWN = [
  "# Service Agreement",
  "This agreement is made between the parties named below.",
  "The Supplier delivers the goods on time and in good order.",
  "The Buyer pays each invoice within thirty days.",
  "Signed in two copies.",
].join("\n\n");

/** Headings and paragraphs; no numbering part at all. */
export const plainDocument = (): Promise<Uint8Array> => packDocument(fromMarkdown(PLAIN_MARKDOWN));

/** A bulleted, a numbered and a nested list (the Markdown numbering part). */
export const listDocument = (): Promise<Uint8Array> =>
  packDocument(
    fromMarkdown(
      [
        "# Delivery Terms",
        "The following apply to every order.",
        "- Goods are packed securely\n- Goods are insured in transit",
        "Payment happens in stages.",
        "1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection",
        "Closing remarks.",
      ].join("\n\n"),
    ),
  );

/**
 * The integrator's numbered-contract recipe (#1092, #1093, #1094): `Heading 2`
 * and `Heading 3` numbered through their styles' `w:numPr` from one
 * multilevel instance, with body paragraphs between them.
 */
export const styleNumberedDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    [
      "# Agreement",
      "## Scope",
      "### Definitions",
      "The Supplier delivers the goods.",
      "The Buyer pays on delivery.",
      "## Payment",
      "Payment is due in ten days (see clause 1).",
      "Closing paragraph.",
    ].join("\n\n"),
  );
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        multiLevelType: "multilevel",
        levels: [decimalLevel(0, "%1."), decimalLevel(1, "%1.%2.")],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const heading2 = findStyle(document, "Heading2");
  heading2.pPr = { ...heading2.pPr, numPr: paragraphNumberingFromSlots({ numId: 5, ilvl: 0 }) };
  const heading3 = findStyle(document, "Heading3");
  heading3.pPr = { ...heading3.pPr, numPr: paragraphNumberingFromSlots({ numId: 5, ilvl: 1 }) };
  return packDocument(document);
};

/** A clause numbered by a direct `w:numPr` next to plain body text. */
export const directNumberedDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    ["# Terms", "Numbered clause one", "Unnumbered body text.", "Numbered clause two"].join("\n\n"),
  );
  document.package.numbering = {
    abstractNums: [
      { abstractNumId: 7, multiLevelType: "singleLevel", levels: [decimalLevel(0, "(%1)")] },
    ],
    nums: [{ numId: 7, abstractNumId: 7 }],
  };
  for (const text of ["Numbered clause one", "Numbered clause two"]) {
    const target = findParagraph(document, text);
    target.formatting = {
      ...target.formatting,
      numPr: paragraphNumberingFromSlots({ numId: 7, ilvl: 0 }),
    };
  }
  return packDocument(document);
};

/** One decimal definition that no paragraph uses (instance 901, #1103). */
export const UNUSED_NUMBERING: Numbering = {
  abstractNums: [
    { abstractNumId: 901, multiLevelType: "multilevel", levels: [decimalLevel(0, "%1.")] },
  ],
  nums: [{ numId: 901, abstractNumId: 901 }],
};

/** Plain prose plus a numbering part none of it uses. */
export const unusedNumberingDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(PLAIN_MARKDOWN);
  document.package.numbering = structuredClone(UNUSED_NUMBERING);
  return packDocument(document);
};

/** A Markdown table between paragraphs, and a built table with a header row. */
export const tableDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    [
      "# Price Schedule",
      "The prices below apply.",
      "| Item | Price |\n| --- | --- |\n| Widget | 10 |\n| Gadget | 20 |",
      "Taxes are extra.",
    ].join("\n\n"),
  );
  document.package.document.content.push(
    table({
      header: ["Milestone", "Date"],
      rows: [
        ["Kick-off", "May"],
        ["Delivery", "June"],
      ],
    }),
    paragraph("Schedules may change by agreement."),
  );
  return packDocument(document);
};

/** A footnote and an endnote referenced from body paragraphs. */
export const notesDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(PLAIN_MARKDOWN);
  const content = document.package.document.content;
  const footnoteId = 1;
  document.package.footnotes = [
    { type: "footnote", id: footnoteId, content: [paragraph("As defined in the order form.")] },
  ];
  content.push(
    paragraph([
      run("Goods are defined in the order form."),
      { type: "run", content: [{ type: "footnoteRef", id: footnoteId }] },
    ]),
    paragraph([run("Warranty terms apply."), endnote(document, "See the warranty schedule.")]),
  );
  return packDocument(document);
};

/** Open comments, a reply and a resolved thread, saved through the reviewer. */
export const commentDocument = async (): Promise<Uint8Array> => {
  const reviewer = await openReviewer(await plainDocument());
  applyOrThrow(reviewer, {
    mode: "direct",
    operations: [
      {
        id: "c1",
        type: "commentOnBlock",
        blockId: blockIdOf(reviewer, "The Supplier delivers the goods on time and in good order."),
        comment: { text: "Define good order." },
      },
      {
        id: "c2",
        type: "commentOnBlock",
        blockId: blockIdOf(reviewer, "The Buyer pays each invoice within thirty days."),
        comment: { text: "Thirty days is long." },
      },
    ],
  });
  const comment = (text: string) => {
    const found = reviewer.getComments().find((candidate) => candidate.text === text);
    if (!found) {
      throw new Error(`fixture comment "${text}" was not created`);
    }
    return found;
  };
  reviewer.replyTo(comment("Define good order."), { text: "Good order means undamaged." });
  reviewer.resolveComment(String(comment("Thirty days is long.").id));
  return new Uint8Array(await reviewer.toBuffer());
};

/** Pending tracked insertions, deletions and a replaced word. */
export const trackedChangesDocument = async (): Promise<Uint8Array> => {
  const reviewer = await openReviewer(await listDocument(), "Earlier Reviewer");
  applyOrThrow(reviewer, {
    mode: "tracked-changes",
    operations: [
      {
        id: "t1",
        type: "replaceInBlock",
        blockId: blockIdOf(reviewer, "The following apply to every order."),
        find: "every",
        replace: "each",
      },
      {
        id: "t2",
        type: "insertAfterBlock",
        blockId: blockIdOf(reviewer, "Payment happens in stages."),
        text: "Stages are invoiced separately.",
      },
      {
        id: "t3",
        type: "deleteBlock",
        blockId: blockIdOf(reviewer, "Goods are insured in transit"),
      },
    ],
  });
  return new Uint8Array(await reviewer.toBuffer());
};

/** Every fixture, by name; scenarios iterate this. */
export const FIXTURES = {
  plain: plainDocument,
  lists: listDocument,
  styleNumbered: styleNumberedDocument,
  directNumbered: directNumberedDocument,
  unusedNumbering: unusedNumberingDocument,
  tables: tableDocument,
  notes: notesDocument,
  comments: commentDocument,
  trackedChanges: trackedChangesDocument,
} as const satisfies Record<string, () => Promise<Uint8Array>>;

export type FixtureName = keyof typeof FIXTURES;
export const FIXTURE_NAMES = Object.keys(FIXTURES) as FixtureName[];
