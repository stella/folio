/**
 * Cross-reader consistency: every surface that reads a document to a person or
 * a model must agree on each block's kind, heading level and the number a
 * reader sees beside it. Each reader has its own tests, and each passed while
 * two of them disagreed about the same numbered heading (`## Scope` in the
 * Markdown, `1.` in `getContent()`; `listItem` in the snapshot, a heading
 * everywhere else), because no test read one document through more than one
 * of them.
 *
 * Agreement on a freshly opened document is not enough: the readers also
 * agreed on every opened document while their labels stayed at the numbers
 * read at open through every later insert or delete. The same comparison
 * therefore runs on the live reviewer after each step of seeded random
 * sequences of list edits, and against the package each step saves.
 *
 * A reader added later (or a field a reader starts to expose) belongs in
 * `readReviewer` below; a numbering shape a reader got wrong belongs in
 * `buildNumberedDocument`; a list edit, in `LIST_EDITS`.
 */

import { describe, expect, test } from "bun:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown, toMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  docxToMarkdown,
  ensureParaIds,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
  type FolioDocumentOperation,
  isFolioAIContentBlock,
} from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCall } from "./execute";
import type { FolioAgentBlock } from "./types";

type Document = ReturnType<typeof fromMarkdown>;
type Paragraph = Extract<Document["package"]["document"]["content"][number], { type: "paragraph" }>;
type NumberingLevel = NonNullable<
  NonNullable<Document["package"]["numbering"]>["abstractNums"]
>[number]["levels"][number];

/** A decimal level unless the options say otherwise. */
type LevelOptions = Pick<NumberingLevel, "ilvl" | "lvlText"> & Partial<NumberingLevel>;

const level = (options: LevelOptions): NumberingLevel => ({
  start: 1,
  numFmt: "decimal",
  suffix: "space",
  pPr: { indentLeft: 0, indentFirstLine: 0 },
  ...options,
});

/** The ids the fixture's numbering definitions use. */
const CLAUSES = 5;
const ARTICLES = 6;
const HIDDEN = 7;

/**
 * A synthetic numbered contract in the shapes real ones use:
 *
 * - `Heading 2` / `Heading 3` numbered through their styles' `w:numPr`
 *   (`1.`, `1.1.`), with a body paragraph continuing the same instance at the
 *   heading's child level, so a heading must advance the counters the list
 *   items after it read;
 * - a heading numbered by a direct `w:numPr`, and an outline-level paragraph
 *   (no heading style) numbered directly;
 * - a heading and a body paragraph on a level whose marker is hidden
 *   (`w:vanish`), and a `Heading 2` and a body paragraph that cancel
 *   numbering with the reserved `w:numId="0"`: numbered in the package, no
 *   number on the page;
 * - a body paragraph numbered at a level its instance does not define, which
 *   Word paints no marker for;
 * - an ordinary bulleted and an ordinary numbered list.
 */
const buildNumberedDocument = async (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    [
      "# Agreement",
      "## Scope",
      "### Definitions",
      "### Terms",
      "Clause body item",
      "The Supplier delivers the goods.",
      "Direct article",
      "Outline article",
      "Hidden-number heading",
      "Hidden-number paragraph",
      "Under a hidden number",
      "Unnumbered clause",
      "Cancelled item",
      "Undefined level",
      "- a bullet\n- another bullet",
      "1. first step\n2. second step",
      "## Payment",
      "Payment is due in ten days (see clause 1).",
    ].join("\n\n"),
  );
  const pkg = document.package;
  pkg.numbering = {
    abstractNums: [
      ...(pkg.numbering?.abstractNums ?? []),
      {
        abstractNumId: CLAUSES,
        multiLevelType: "multilevel",
        levels: [level({ ilvl: 0, lvlText: "%1." }), level({ ilvl: 1, lvlText: "%1.%2." })],
      },
      {
        abstractNumId: ARTICLES,
        multiLevelType: "multilevel",
        levels: [
          level({ ilvl: 0, lvlText: "(%1)" }),
          level({ ilvl: 1, lvlText: "%2)", numFmt: "lowerLetter" }),
        ],
      },
      {
        abstractNumId: HIDDEN,
        multiLevelType: "multilevel",
        levels: [
          level({ ilvl: 0, lvlText: "Art. %1", rPr: { hidden: true } }),
          level({ ilvl: 1, lvlText: "%1.%2." }),
        ],
      },
    ],
    nums: [
      ...(pkg.numbering?.nums ?? []),
      { numId: CLAUSES, abstractNumId: CLAUSES },
      { numId: ARTICLES, abstractNumId: ARTICLES },
      { numId: HIDDEN, abstractNumId: HIDDEN },
    ],
  };

  const style = (styleId: string) => {
    const found = pkg.styles?.styles.find((candidate) => candidate.styleId === styleId);
    if (!found) {
      throw new Error(`fixture style ${styleId} is missing`);
    }
    return found;
  };
  const heading2 = style("Heading2");
  heading2.pPr = {
    ...heading2.pPr,
    numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 0 }),
  };
  const heading3 = style("Heading3");
  heading3.pPr = {
    ...heading3.pPr,
    numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 1 }),
  };

  const paragraph = (text: string): Paragraph => {
    const found = pkg.document.content.find(
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
  const format = (text: string, formatting: NonNullable<Paragraph["formatting"]>): void => {
    const target = paragraph(text);
    target.formatting = { ...target.formatting, ...formatting };
  };
  format("Clause body item", { numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 1 }) });
  format("Direct article", {
    styleId: "Heading4",
    numPr: paragraphNumberingFromSlots({ numId: ARTICLES, ilvl: 0 }),
  });
  format("Outline article", {
    outlineLevel: { kind: "heading", level: 3 },
    numPr: paragraphNumberingFromSlots({ numId: ARTICLES, ilvl: 1 }),
  });
  format("Hidden-number heading", {
    styleId: "Heading4",
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 0 }),
  });
  format("Hidden-number paragraph", {
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 0 }),
  });
  format("Under a hidden number", {
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 1 }),
  });
  format("Unnumbered clause", {
    styleId: "Heading2",
    numPr: paragraphNumberingFromSlots({ numId: 0 }),
  });
  format("Cancelled item", { numPr: paragraphNumberingFromSlots({ numId: 0 }) });
  format("Undefined level", { numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 5 }) });

  return (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;
};

type Kind = "heading" | "listItem" | "paragraph";

/** What one reader says about one block. */
type BlockView = {
  text: string;
  kind: Kind;
  /** One-based heading level; absent when the reader shows no heading. */
  headingLevel?: number;
  /** The number or bullet a reader sees beside the text; absent when none. */
  number?: string;
};

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/** A bullet reads as `-` in Markdown and as its glyph elsewhere; both are "a bullet". */
const BULLET = "(bullet)";
const normalizeNumber = (marker: string | undefined): string | undefined => {
  if (marker === undefined || marker.length === 0) {
    return undefined;
  }
  return /[\p{L}\p{N}]/u.test(marker) ? marker : BULLET;
};

/**
 * `getContent()` / the AI snapshot. A heading's `displayLabel` is its style
 * id when it carries no number, so only a label that is not the style id is a
 * number a reader sees.
 */
const snapshotView = (block: {
  text: string;
  kind: string;
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
}): BlockView => {
  const number =
    block.displayLabel !== undefined && block.displayLabel !== block.styleId
      ? normalizeNumber(block.displayLabel)
      : undefined;
  return blockView({
    text: block.text,
    kind: asKind(block.kind),
    headingLevel: block.headingLevel,
    number,
  });
};

const asKind = (kind: string): Kind => {
  if (kind === "heading" || kind === "listItem" || kind === "paragraph") {
    return kind;
  }
  throw new Error(`unexpected block kind ${kind}`);
};

/** A line with no heading tag reads as a list item when it shows a number. */
const bodyView = (text: string, number: string | undefined): BlockView =>
  blockView({
    text,
    kind: number === undefined ? "paragraph" : "listItem",
    headingLevel: undefined,
    number,
  });

/**
 * `docxToMarkdown`, one non-blank line per block (the fixture has no
 * multi-line paragraphs), matched to the blocks in document order.
 */
const markdownViews = (markdown: string, texts: readonly string[]): BlockView[] =>
  linesFor({
    output: markdown,
    texts,
    view: (prefix, text) => {
      const heading = /^(?<hashes>#{1,6})(?:\s+(?<marker>.*))?$/u.exec(prefix);
      if (heading?.groups) {
        return blockView({
          text,
          kind: "heading",
          headingLevel: heading.groups["hashes"]?.length,
          number: normalizeNumber(heading.groups["marker"]),
        });
      }
      return bodyView(text, normalizeNumber(prefix === "-" ? "•" : prefix));
    },
  });

/** `getContentAsText()`: `[id] (h2) 1. Scope`, `[id] a) item`, `[id] text`. */
const contentTextViews = (lines: string, texts: readonly string[]): BlockView[] =>
  linesFor({
    output: lines,
    texts,
    view: (prefix, text) => {
      const line = /^\[[^\]]+\](?:\s+\(h(?<level>\d)\))?(?:\s+(?<marker>.*))?$/u.exec(prefix);
      if (!line?.groups) {
        throw new Error(`unexpected getContentAsText line prefix "${prefix}"`);
      }
      const number = normalizeNumber(line.groups["marker"]);
      const headingLevel = line.groups["level"];
      if (headingLevel !== undefined) {
        return blockView({ text, kind: "heading", headingLevel: Number(headingLevel), number });
      }
      return bodyView(text, number);
    },
  });

/** One view per non-blank output line, each line ending in its block's text. */
type LinesForOptions = {
  output: string;
  texts: readonly string[];
  view: (prefix: string, text: string) => BlockView;
};

const linesFor = ({ output, texts, view }: LinesForOptions): BlockView[] => {
  const lines = output.split("\n").filter((line) => line.trim().length > 0);
  expect(lines).toHaveLength(texts.length);
  return lines.map((line, index) => {
    const text = texts[index] ?? "";
    expect(line.endsWith(text)).toBe(true);
    return view(line.slice(0, line.length - text.length).trim(), text);
  });
};

type BlockViewOptions = {
  text: string;
  kind: Kind;
  headingLevel: number | undefined;
  number: string | undefined;
};

/** A view without the fields the reader left undefined. */
const blockView = ({ text, kind, headingLevel, number }: BlockViewOptions): BlockView => {
  const view: BlockView = { text, kind };
  if (headingLevel !== undefined) {
    view.headingLevel = headingLevel;
  }
  if (number !== undefined) {
    view.number = number;
  }
  return view;
};

/** The `read_document` row fields that restate a snapshot block's. */
const rowFields = (row: FolioAgentBlock) => ({
  blockId: row.blockId,
  kind: row.kind,
  displayLabel: row.displayLabel,
  headingLevel: row.headingLevel,
  listLevel: row.listLevel,
});

const MARKDOWN_OPTIONS = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  footnotes: "keep",
} as const;

/** Every reader of the package `bytes` holds, opened afresh. */
const readAll = async (bytes: Uint8Array) =>
  readReviewer(
    await FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes)),
    await docxToMarkdown(toArrayBuffer(bytes), MARKDOWN_OPTIONS),
  );

/** Every reader of a reviewer as it stands, Markdown from its current document. */
const readLive = (reviewer: FolioDocxReviewer) =>
  readReviewer(reviewer, toMarkdown(reviewer.toDocument(), MARKDOWN_OPTIONS));

const readReviewer = (reviewer: FolioDocxReviewer, markdown: string) => {
  const bridge = createReviewerBridge(reviewer);
  const content = reviewer.getContent().filter(isFolioAIContentBlock);
  const snapshot = bridge.snapshot().blocks.filter(isFolioAIContentBlock);
  const read = executeFolioToolCall("read_document", {}, bridge);
  if (!read.ok) {
    throw new Error(read.error);
  }
  const texts = content.map(({ text }) => text);
  const styleIds = new Map(content.map((block) => [block.id, block.styleId]));
  return {
    content,
    rows: read.result,
    views: {
      getContent: content.map(snapshotView),
      snapshot: snapshot.map(snapshotView),
      readDocument: read.result.map((row) =>
        snapshotView({ ...row, styleId: styleIds.get(row.blockId) }),
      ),
      getContentAsText: contentTextViews(reviewer.getContentAsText(), texts),
      markdown: markdownViews(markdown, texts),
    },
  };
};

type Block = ReturnType<FolioDocxReviewer["getContent"]>[number];
type Mode = "direct" | "tracked-changes";

/** A small seeded generator (mulberry32), so a failing sequence replays from its seed. */
const seeded = (seed: number) => {
  let state = seed >>> 0;
  const next = (): number => {
    state = (state + 0x6d_2b_79_f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4_294_967_296;
  };
  const pick = <T>(items: readonly T[]): T | undefined => items[Math.floor(next() * items.length)];
  return { next, pick };
};

type Random = ReturnType<typeof seeded>;

/** One list edit: the operations of one batch, or `null` when the document offers no target. */
type ListEdit = (
  blocks: readonly Block[],
  random: Random,
  step: number,
) => FolioDocumentOperation[] | null;

const listed = (blocks: readonly Block[]): Block[] =>
  blocks.filter((block) => block.listReference !== undefined);

/** The edits that change which items a list has, their levels, or where it restarts. */
const LIST_EDITS: Record<string, ListEdit> = {
  insertAfterItem: (blocks, random, step) => {
    const anchor = random.pick(listed(blocks)) ?? random.pick(blocks);
    return anchor
      ? [{ id: "e", type: "insertAfterBlock", blockId: anchor.id, text: `Inserted ${step}` }]
      : null;
  },
  // A style of its own: the item keeps the anchor's numbering reference.
  insertStyledItem: (blocks, random, step) => {
    const anchor = random.pick(listed(blocks)) ?? random.pick(blocks);
    return anchor
      ? [
          {
            id: "e",
            type: "insertAfterBlock",
            blockId: anchor.id,
            text: `Styled ${step}`,
            styleId: "Normal",
          },
        ]
      : null;
  },
  // Into or out of a style that numbers and outlines its paragraphs.
  restyle: (blocks, random) => {
    const target = random.pick(blocks);
    return target
      ? [
          {
            id: "e",
            type: "setBlockParagraphProperties",
            blockId: target.id,
            properties: { styleId: random.pick(["Heading2", "Heading3", null]) ?? null },
          },
        ]
      : null;
  },
  insertIntoList: (blocks, random, step) => {
    const anchor = random.pick(blocks);
    const member = random.pick(listed(blocks))?.listReference;
    return anchor && member
      ? [
          {
            id: "e",
            type: "insertBeforeBlock",
            blockId: anchor.id,
            text: `Joined ${step}`,
            numbering: { numId: member.numId, level: member.level },
          },
        ]
      : null;
  },
  deleteItem: (blocks, random) => {
    const target = random.pick(listed(blocks)) ?? random.pick(blocks);
    return target && blocks.length > 6
      ? [{ id: "e", type: "deleteBlock", blockId: target.id }]
      : null;
  },
  moveItem: (blocks, random) => {
    const item = random.pick(listed(blocks));
    const to = random.pick(blocks.filter((block) => block.id !== item?.id));
    const reference = item?.listReference;
    return item && to && reference
      ? [
          {
            id: "copy",
            type: "insertAfterBlock",
            blockId: to.id,
            text: item.text,
            numbering: { numId: reference.numId, level: reference.level },
          },
          { id: "remove", type: "deleteBlock", blockId: item.id },
        ]
      : null;
  },
  // Up or down a level, sometimes to one the list does not define.
  changeLevel: (blocks, random) => {
    const item = random.pick(listed(blocks));
    const reference = item?.listReference;
    if (!item || !reference) return null;
    const nextLevel = Math.max(0, Math.min(8, reference.level + (random.next() < 0.5 ? -1 : 1)));
    return [
      {
        id: "e",
        type: "setBlockParagraphProperties",
        blockId: item.id,
        properties: { numbering: { numId: reference.numId, level: nextLevel } },
      },
    ];
  },
  restartNumbering: (blocks, random) => {
    const target = random.pick(blocks);
    return target
      ? [
          {
            id: "e",
            type: "setBlockParagraphProperties",
            blockId: target.id,
            properties: { numbering: { start: "new", kind: "numbered" } },
          },
        ]
      : null;
  },
  removeFromList: (blocks, random) => {
    const item = random.pick(listed(blocks));
    return item
      ? [
          {
            id: "e",
            type: "setBlockParagraphProperties",
            blockId: item.id,
            properties: { numbering: null },
          },
        ]
      : null;
  },
};

type Views = Awaited<ReturnType<typeof readAll>>["views"];

/**
 * Every reader agrees with `getContent()`. Clean Markdown shows a pending
 * change resolved, so it is left out while changes are pending.
 */
const expectReadersAgree = (views: Views, withMarkdown: boolean): void => {
  for (const [reader, view] of Object.entries(views)) {
    if (reader === "markdown" && !withMarkdown) continue;
    expect({ reader, view }).toEqual({ reader, view: views.getContent });
  }
};

/**
 * The live reviewer's readers agree with each other, its rows restate its
 * labels, and the package it saves reads the same through every reader.
 */
const expectLiveAndSavedAgree = async (reviewer: FolioDocxReviewer): Promise<void> => {
  const withMarkdown = reviewer.getChanges().length === 0;
  const live = readLive(reviewer);
  expectReadersAgree(live.views, withMarkdown);
  expect(live.rows.map(rowFields)).toEqual(
    live.content.map((block) => ({
      blockId: block.id,
      kind: block.kind,
      displayLabel: block.displayLabel,
      headingLevel: block.headingLevel,
      listLevel: block.listLevel,
    })),
  );

  const saved = await readAll(new Uint8Array(await reviewer.toBuffer()));
  expect({ saved: saved.views.getContent }).toEqual({ saved: live.views.getContent });
  expectReadersAgree(saved.views, withMarkdown);
};

/**
 * Apply `steps` random list edits and, after each, compare every reader of the
 * live reviewer with each other, with the package the step saves, and that
 * package's readers with each other.
 */
const runListEdits = async (seed: number, mode: Mode, steps: number): Promise<void> => {
  const random = seeded(seed);
  const reviewer = await FolioDocxReviewer.fromBuffer(
    toArrayBuffer(await buildNumberedDocument()),
    { author: "Agent" },
  );
  const log: string[] = [];
  const check = async (body: () => Promise<void> | void): Promise<void> => {
    try {
      await body();
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`seed ${seed} (${mode}) after ${log.join(" → ")}\n${detail}`, {
        cause: error,
      });
    }
  };
  // What rejecting every pending change must give back: the document as it
  // stood before the changes still pending.
  let settled = readLive(reviewer).views.getContent;
  // Tracked runs resolve their changes now and then, and at the end.
  const resolve = async (resolution: "acceptAll" | "rejectAll"): Promise<void> => {
    reviewer[resolution]();
    log.push(resolution);
    await check(async () => {
      if (resolution === "rejectAll") {
        expect(readLive(reviewer).views.getContent).toEqual(settled);
      }
      settled = readLive(reviewer).views.getContent;
      await expectLiveAndSavedAgree(reviewer);
    });
  };
  for (let step = 0; step < steps; step += 1) {
    if (mode === "tracked-changes" && random.next() < 0.2) {
      await resolve(random.next() < 0.5 ? "acceptAll" : "rejectAll");
      continue;
    }
    const name = random.pick(Object.keys(LIST_EDITS)) ?? "insertAfterItem";
    const blocks = reviewer.getContent().filter(isFolioAIContentBlock);
    const operations = LIST_EDITS[name]?.(blocks, random, step);
    if (!operations) continue;
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations,
    });
    log.push(`${name}:${result.status}`);
    await check(() => expectLiveAndSavedAgree(reviewer));
  }
  if (mode === "tracked-changes") {
    await resolve("rejectAll");
  }
};

describe("readers agree after list edits", () => {
  for (const seed of [1, 2, 3, 5, 8, 13]) {
    test(`seed ${seed}: direct edits`, async () => {
      await runListEdits(seed, "direct", 8);
    });
  }
  for (const seed of [21, 34, 55, 89, 144]) {
    test(`seed ${seed}: tracked edits`, async () => {
      await runListEdits(seed, "tracked-changes", 8);
    });
  }
});

describe("readers agree on numbered headings and lists", () => {
  test("every reader shows each block's kind, heading level and number alike", async () => {
    const { views } = await readAll(await buildNumberedDocument());

    // The fixture's intent, pinned once so a reader that drifts shows as a
    // disagreement below rather than rewriting the expectation.
    expect(views.getContent).toEqual([
      { text: "Agreement", kind: "heading", headingLevel: 1 },
      { text: "Scope", kind: "heading", headingLevel: 2, number: "1." },
      { text: "Definitions", kind: "heading", headingLevel: 3, number: "1.1." },
      { text: "Terms", kind: "heading", headingLevel: 3, number: "1.2." },
      { text: "Clause body item", kind: "listItem", number: "1.3." },
      { text: "The Supplier delivers the goods.", kind: "paragraph" },
      { text: "Direct article", kind: "heading", headingLevel: 4, number: "(1)" },
      { text: "Outline article", kind: "heading", headingLevel: 4, number: "a)" },
      { text: "Hidden-number heading", kind: "heading", headingLevel: 4 },
      { text: "Hidden-number paragraph", kind: "paragraph" },
      // A hidden marker still counts: the child shows the parent's number.
      { text: "Under a hidden number", kind: "listItem", number: "2.1." },
      { text: "Unnumbered clause", kind: "heading", headingLevel: 2 },
      { text: "Cancelled item", kind: "paragraph" },
      // A level the instance does not define shows no marker: prose.
      { text: "Undefined level", kind: "paragraph" },
      { text: "a bullet", kind: "listItem", number: BULLET },
      { text: "another bullet", kind: "listItem", number: BULLET },
      { text: "first step", kind: "listItem", number: "1." },
      { text: "second step", kind: "listItem", number: "2." },
      { text: "Payment", kind: "heading", headingLevel: 2, number: "2." },
      { text: "Payment is due in ten days (see clause 1).", kind: "paragraph" },
    ]);
    for (const [reader, view] of Object.entries(views)) {
      expect({ reader, view }).toEqual({ reader, view: views.getContent });
    }
  });

  test("read_document rows restate the snapshot's label and levels", async () => {
    const { content, rows } = await readAll(await buildNumberedDocument());

    expect(rows.map(rowFields)).toEqual(
      content.map((block) => ({
        blockId: block.id,
        kind: block.kind,
        displayLabel: block.displayLabel,
        headingLevel: block.headingLevel,
        listLevel: block.listLevel,
      })),
    );
    // Rows stay compact: a field the block lacks is absent, not `undefined`.
    const body = rows.find(({ text }) => text === "The Supplier delivers the goods.");
    expect(body && Object.keys(body).sort()).toEqual(["blockId", "blockTextHash", "kind", "text"]);
    expect(rows.find(({ text }) => text === "Definitions")).toMatchObject({
      kind: "heading",
      displayLabel: "1.1.",
      headingLevel: 3,
      listLevel: 1,
    });
  });
});
