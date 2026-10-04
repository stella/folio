/**
 * Synthetic documents built the way an integrator builds them: Markdown
 * through `fromMarkdown`, a few model edits through the published builders,
 * `createDocx`, then `ensureParaIds`. Each fixture returns package bytes; the
 * edits that need a reviewer (comments, tracked changes) go through the public
 * operation contract and a save, so a fixture is itself a first round trip.
 */

import { createHash } from "node:crypto";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  endnote,
  ensureParaIds,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
  hyperlink,
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

/** A paraId `ensureParaIds` keeps: eight hex digits, below 0x80000000, not all zero. */
const fixtureParaId = (seed: string): string => {
  const value = createHash("sha256").update(seed).digest().readUInt32BE(0) & 0x7f_ff_ff_ff;
  return value.toString(16).toUpperCase().padStart(8, "0");
};

const isParagraph = (value: object): value is Paragraph =>
  "type" in value && value.type === "paragraph";

/**
 * The paraIds each fixture's model carried into `createDocx`, keyed by the
 * exact bytes a fixture returned; per package, so no answer depends on what
 * else the process built.
 */
const PINNED_PARA_IDS = new WeakMap<Uint8Array, ReadonlySet<string>>();

/**
 * The paraIds the model behind `bytes` carried, rather than ids
 * `ensureParaIds` minted from the serialized bytes (a note separator) or a
 * reviewer edit minted while the fixture was built. `bytes` must be the
 * object `packFixture` or a fixture returned (a reviewer-built fixture
 * carries the ids of the fixture it opened).
 */
export const pinnedParaIdsOf = (bytes: Uint8Array): ReadonlySet<string> => {
  const pinned = PINNED_PARA_IDS.get(bytes);
  if (pinned === undefined) {
    throw new TypeError("pinnedParaIdsOf: these bytes are not a fixture packFixture returned");
  }
  return pinned;
};

/**
 * A copy of `document` with every story paragraph given a `w14:paraId`
 * named by its story and its position there (depth first, so table cells
 * and text boxes count too), and the ids it carries. The reviewer's block
 * id is the paragraph's paraId, and flows pin fixture blocks by id
 * (support/flow-file.ts); an id `ensureParaIds` mints derives from a hash of
 * document.xml as the serializer writes it, so a serializer change would
 * re-mint every id and detach every flow.
 */
const pinParaIds = (source: FolioDocument): { document: FolioDocument; pinned: Set<string> } => {
  const document = structuredClone(source);
  const pkg = document.package;
  const stories: [string, unknown][] = [
    ["body", pkg.document.content],
    ...[...(pkg.headers ?? [])].map(([rId, part]): [string, unknown] => [`header:${rId}`, part]),
    ...[...(pkg.footers ?? [])].map(([rId, part]): [string, unknown] => [`footer:${rId}`, part]),
    ...(pkg.footnotes ?? []).map((note): [string, unknown] => [`footnote:${note.id}`, note]),
    ...(pkg.endnotes ?? []).map((note): [string, unknown] => [`endnote:${note.id}`, note]),
  ];
  const pinned = new Set<string>();
  for (const [story, root] of stories) {
    let ordinal = 0;
    const visit = (value: unknown): void => {
      if (Array.isArray(value)) {
        for (const item of value) visit(item);
        return;
      }
      if (typeof value !== "object" || value === null) return;
      if (isParagraph(value)) {
        if (value.paraId === undefined) {
          const seed = `${story}:${ordinal}`;
          let id = fixtureParaId(seed);
          for (let salt = 1; pinned.has(id) || id === "00000000"; salt += 1) {
            id = fixtureParaId(`${seed}:${salt}`);
          }
          value.paraId = id;
        }
        pinned.add(value.paraId);
        ordinal += 1;
      }
      for (const item of Object.values(value)) visit(item);
    };
    visit(root);
  }
  return { document, pinned };
};

/**
 * `createDocx` + `ensureParaIds`, the recipe the published README gives:
 * `ensureParaIds` mints every paraId. For an ad-hoc document; a fixture a
 * flow can name blocks of goes through `packFixture`.
 */
export const packDocument = async (document: FolioDocument): Promise<Uint8Array> =>
  (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;

/**
 * `packDocument` on a copy of `source` whose story paragraphs carry pinned
 * paraIds (see `pinParaIds`), so `ensureParaIds` mints only the note
 * separators the serializer adds; `pinnedParaIdsOf` answers for the
 * returned bytes. Positional ids repeat across fixtures, so two unrelated
 * fixtures share paraIds: never compare or merge them by paragraph identity.
 */
export const packFixture = async (source: FolioDocument): Promise<Uint8Array> => {
  const { document, pinned } = pinParaIds(source);
  const bytes = await packDocument(document);
  PINNED_PARA_IDS.set(bytes, pinned);
  return bytes;
};

/** A reviewer-built fixture's save, carrying the pinned ids of the fixture it opened. */
const saveFixture = async (reviewer: Reviewer, base: Uint8Array): Promise<Uint8Array> => {
  const bytes = new Uint8Array(await reviewer.toBuffer());
  PINNED_PARA_IDS.set(bytes, pinnedParaIdsOf(base));
  return bytes;
};

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
export const plainDocument = (): Promise<Uint8Array> => packFixture(fromMarkdown(PLAIN_MARKDOWN));

/** A bulleted, a numbered and a nested list (the Markdown numbering part). */
export const listDocument = (): Promise<Uint8Array> =>
  packFixture(
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
  return packFixture(document);
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
  return packFixture(document);
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
  return packFixture(document);
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
  return packFixture(document);
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
  return packFixture(document);
};

/** Open comments, a reply and a resolved thread, saved through the reviewer. */
export const commentDocument = async (): Promise<Uint8Array> => {
  const base = await plainDocument();
  const reviewer = await openReviewer(base);
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
  return saveFixture(reviewer, base);
};

/** Pending tracked insertions, deletions and a replaced word. */
export const trackedChangesDocument = async (): Promise<Uint8Array> => {
  const base = await listDocument();
  const reviewer = await openReviewer(base, "Earlier Reviewer");
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
  return saveFixture(reviewer, base);
};

/** Emoji outside the Basic Multilingual Plane, a modifier and a ZWJ sequence, next to words. */
export const EMOJI_TEXTS = [
  "Approved 👍🏽 by the family 👨‍👩‍👧 today.",
  "Launch 🚀",
  "🎉 opens the party.",
] as const;

/** Paragraphs whose words sit next to surrogate pairs. */
export const emojiDocument = (): Promise<Uint8Array> =>
  packFixture(fromMarkdown(["# Emoji", ...EMOJI_TEXTS].join("\n\n")));

/** A table whose first column merges vertically through its last two rows. */
export const mergedTableDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown("# Merged cells\n\nThe table below merges a column.");
  document.package.document.content.push(
    table({
      header: ["Party", "Role", "Share"],
      rows: [
        [{ content: [paragraph("Supplier")], vMerge: "restart" }, "Delivers", "60"],
        [{ content: [paragraph([])], vMerge: "continue" }, "Installs", "40"],
      ],
    }),
    paragraph("After the table."),
  );
  return packFixture(document);
};

type HeaderFooterKind = "default" | "first" | "even";

/** A header or footer part holding one paragraph; a footer's ends in a PAGE field. */
const headerFooterPart = (
  type: "header" | "footer",
  hdrFtrType: HeaderFooterKind,
  text: string,
) => ({
  type,
  hdrFtrType,
  content: [
    paragraph([
      run(text),
      ...(type === "footer"
        ? [
            {
              type: "simpleField" as const,
              instruction: "PAGE",
              fieldType: "PAGE" as const,
              content: [run("1")],
            },
          ]
        : []),
    ]),
  ],
});

/** The text box's own paragraph in `storiesDocument`. */
export const TEXT_BOX_TEXT = "Boxed note on delivery terms.";

/**
 * Every story a reviewer can reach beyond the body, and the body features
 * edits collide with: a footnote and an endnote, default / first-page /
 * even-page headers and footers (a PAGE field in each footer), two sections
 * (the first ends in a paragraph carrying its `w:sectPr`), a paragraph that
 * holds a text box, a hyperlink, a date field, a run-formatting boundary, a
 * list, an emoji, and a comment whose range spans a whole table.
 */
export const storiesDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    [
      "# Master Agreement",
      "The parties agree as follows.",
      "- Goods are listed in the schedule\n- Prices are fixed for a year",
      "Payment is due on receipt.",
    ].join("\n\n"),
  );
  const pkg = document.package;
  pkg.headers = new Map([
    ["rIdStoryHeader1", headerFooterPart("header", "default", "Master Agreement header")],
    ["rIdStoryHeader2", headerFooterPart("header", "first", "First page header")],
    ["rIdStoryHeader3", headerFooterPart("header", "even", "Even page header")],
  ]);
  pkg.footers = new Map([
    ["rIdStoryFooter1", headerFooterPart("footer", "default", "Page ")],
    ["rIdStoryFooter2", headerFooterPart("footer", "even", "Even page ")],
  ]);
  pkg.settings = { defaultTabStop: 720, ...pkg.settings, evenAndOddHeaders: true };
  const references = {
    headerReferences: [
      { type: "default" as const, rId: "rIdStoryHeader1" },
      { type: "first" as const, rId: "rIdStoryHeader2" },
      { type: "even" as const, rId: "rIdStoryHeader3" },
    ],
    footerReferences: [
      { type: "default" as const, rId: "rIdStoryFooter1" },
      { type: "even" as const, rId: "rIdStoryFooter2" },
    ],
    titlePg: true,
  };
  pkg.footnotes = [
    { type: "footnote", id: 1, content: [paragraph("Schedules are part of this agreement.")] },
  ];
  const textBox = {
    type: "shape" as const,
    shape: {
      type: "shape" as const,
      shapeType: "textBox" as const,
      id: "story-text-box",
      size: { width: 2_000_000, height: 800_000 },
      textBody: { content: [paragraph(TEXT_BOX_TEXT)] },
    },
  };
  pkg.document.content.push(
    paragraph([
      run("Delivery follows the schedule"),
      { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
      run(" and the "),
      run("warranty", { bold: true }),
      run(" applies"),
      endnote(document, "Warranty terms are in annex two."),
      run("."),
    ]),
    paragraph([
      run("The box beside this clause restates it. "),
      { type: "run", content: [textBox] },
    ]),
    paragraph([
      run("Notices go to "),
      hyperlink({ text: "the notice address", href: "https://example.com/notices" }),
      run(" dated "),
      { type: "simpleField", instruction: "DATE", fieldType: "DATE", content: [run("1 May 2026")] },
      run(" in writing."),
    ]),
    { ...paragraph("This clause ends the first section."), sectionProperties: references },
    paragraph([{ type: "commentRangeStart", id: 41 }, run("The schedule below is binding.")]),
    table({
      header: ["Item", "Price"],
      rows: [
        ["Widget", "10"],
        ["Gadget", "20"],
      ],
    }),
    paragraph([
      run("Prices exclude taxes."),
      { type: "commentRangeEnd", id: 41 },
      { type: "commentReference", id: 41 },
    ]),
    paragraph("Signed by both parties 👍🏽 today."),
  );
  pkg.document.comments = [
    { id: 41, author: "Earlier Reviewer", content: [paragraph("Check the whole schedule.")] },
  ];
  pkg.document.finalSectionProperties = {
    ...pkg.document.finalSectionProperties,
    ...references,
  };
  return packFixture(document);
};

/**
 * The fixtures the collision scenarios add to `FIXTURES`. Kept apart so the
 * seeded fuzz flows, which pick from `FIXTURES`, replay as before.
 */
export const COLLISION_FIXTURES = {
  emoji: emojiDocument,
  mergedTable: mergedTableDocument,
} as const satisfies Record<string, () => Promise<Uint8Array>>;

/**
 * Fixtures with headers, footers, notes, sections and text boxes. Kept apart
 * from `FIXTURES` so the legacy flows, which pick from it, replay as before.
 */
export const STORY_FIXTURES = {
  stories: storiesDocument,
} as const satisfies Record<string, () => Promise<Uint8Array>>;

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
