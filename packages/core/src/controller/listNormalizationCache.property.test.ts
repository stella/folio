/**
 * A list normalization that reuses the blocks earlier normalizations left
 * equals a normalization of every block, and normalizes no block it could
 * reuse.
 *
 * The cache normalizes a story again from its first changed block, and after
 * its last changed block only while the counters differ from those the
 * earlier run had at the same block. Story versions made by session edits
 * and by model edits that number, renumber, unnumber, insert, remove and move
 * paragraphs are normalized both ways: the documents, which blocks came back
 * unchanged, the inverse operations, and the facts kept for each paragraph
 * must agree, and the cache must normalize exactly the changed blocks and
 * those the counters before them differ at.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import {
  OP_STORIES,
  paragraphPropertySnapshotToken,
  registerParagraphPropertySnapshot,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import {
  computeListMarker,
  type ComputeListMarkerOptions,
  snapshotListCounters,
} from "../docx/listMarkerComputation";
import {
  ListNormalizationCache,
  normalizeListModelRendering,
} from "../docx/listModelNormalization";
import {
  computeListRendering,
  getCachedNumberingMap,
} from "../docx/numberingParser";
import {
  carryParagraphPropertySource,
  paragraphPropertySourceState,
} from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type {
  BlockContent,
  Document,
  ListLevel,
  Paragraph,
} from "../types/document";
import {
  applyEdit,
  editArbitrary,
  editModel,
  modelEditArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";

setDefaultTimeout(propertyTestTimeout(240_000));

const mainBlocks = (document: Document): readonly BlockContent[] =>
  document.package.document.content;

/** The facts kept for a paragraph outside its fields, and its snapshot registration. */
const keptFacts = (block: BlockContent): readonly unknown[] =>
  block.type === "paragraph"
    ? [
        ...paragraphPropertySourceState(block),
        paragraphPropertySnapshotToken(block),
      ]
    : [];

/** Whether each fact is the very same value or object. */
const sameFacts = (
  left: readonly (readonly unknown[])[],
  right: readonly (readonly unknown[])[],
): boolean =>
  left.length === right.length &&
  left.every(
    (facts, index) =>
      facts.length === right[index]?.length &&
      facts.every((fact, at) => fact === right[index]?.[at]),
  );

/** Facts by value: a block normalized anew holds fresh objects of equal content. */
const factValues = (blocks: readonly BlockContent[]): unknown =>
  JSON.parse(JSON.stringify(blocks.map(keptFacts)));

/**
 * The counters before each block of a normalized story, in a form equal for
 * equal counters: each array named by its cell, only arrays something is
 * bound to, in no particular order. Walked apart from the normalization, on
 * copies, since a paragraph's counters do not depend on its marker text.
 */
const countersBeforeEach = (normalized: Document): string[] => {
  const numbering = normalized.package.numbering;
  const counters: ComputeListMarkerOptions = {
    numbering:
      numbering === undefined ? null : getCachedNumberingMap(numbering),
    listCounters: new Map(),
    abstractCounters: new Map(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
  };
  const sorted = <T extends readonly [number, unknown]>(entries: Iterable<T>) =>
    [...entries].toSorted(([left], [right]) => left - right);
  // Paragraphs a content control holds continue the story's lists.
  const advance = (block: BlockContent): void => {
    if (block.type === "blockSdt" || block.type === "blockCustomXml") {
      for (const child of block.content) advance(child);
      return;
    }
    if (block.type !== "paragraph") return;
    const rendering = block.listRendering;
    computeListMarker(
      rendering === undefined
        ? block
        : {
            ...block,
            listRendering: {
              ...rendering,
              marker: rendering.markerTemplate ?? rendering.marker,
            },
          },
      counters,
    );
  };
  return mainBlocks(normalized).map((block) => {
    const snapshot = snapshotListCounters(counters);
    const before = JSON.stringify({
      previous: snapshot.previousList,
      list: sorted(snapshot.listCounters),
      abstract: sorted(snapshot.abstractCounters),
      restarted: snapshot.restartedNumIds.toSorted((a, b) => a - b),
      cells: sorted(snapshot.cells).map(([cell, values]) => [
        cell,
        values.map(String),
      ]),
    });
    advance(block);
    return before;
  });
};

/** A version as its normalization left it, to judge what the next one may reuse. */
type Normalized = {
  document: Document;
  /** The normalized blocks, each of which normalizes to itself. */
  blocks: readonly BlockContent[];
  /** The snapshot registration each block was left with. */
  tokens: readonly unknown[];
  counters: readonly string[];
};

/**
 * How many blocks of `inputs` a normalization may not reuse after `earlier`:
 * the changed blocks between the longest kept prefix and suffix, and the kept
 * blocks after them up to the first the counters before it are the same at.
 */
const blocksToNormalize = (
  inputs: readonly BlockContent[],
  tokens: readonly unknown[],
  counters: readonly string[],
  earlier: Normalized | undefined,
): number => {
  if (earlier === undefined) return inputs.length;
  const length = inputs.length;
  const old = earlier.blocks;
  const kept = (index: number, oldIndex: number) =>
    inputs[index] === old[oldIndex] &&
    tokens[index] === earlier.tokens[oldIndex];
  let prefix = 0;
  while (prefix < length && prefix < old.length && kept(prefix, prefix))
    prefix += 1;
  let suffix = 0;
  while (
    suffix < length - prefix &&
    suffix < old.length - prefix &&
    kept(length - 1 - suffix, old.length - 1 - suffix)
  )
    suffix += 1;
  const shift = length - old.length;
  let reached = length - suffix;
  while (
    reached < length &&
    counters[reached] !== earlier.counters[reached - shift]
  )
    reached += 1;
  return reached - prefix;
};

/**
 * Normalize `version` with the cache, then in full over the same records,
 * and require the same outcome, with no block normalized that the cache
 * could reuse after `earlier`; the cached result is the next version's base.
 */
const normalizedBothWays = (
  version: Document,
  cache: ListNormalizationCache,
  earlier?: Normalized,
): Normalized => {
  const inputs = [...mainBlocks(version)];
  const tokensBefore = inputs.map((block) =>
    block.type === "paragraph"
      ? paragraphPropertySnapshotToken(block)
      : undefined,
  );
  const cached = normalizeListModelRendering(version, { cache });
  const cachedBlocks = mainBlocks(cached.document);
  const cachedFacts = inputs.map(keptFacts);
  const cachedOutputFacts = factValues(cachedBlocks);
  const full = normalizeListModelRendering(version);
  const fullBlocks = mainBlocks(full.document);
  expect(JSON.parse(JSON.stringify(cached.document.package))).toEqual(
    JSON.parse(JSON.stringify(full.document.package)),
  );
  expect(cached.inverse).toEqual(full.inverse);
  // The same blocks come back unchanged.
  expect(cachedBlocks.map((block, index) => block === inputs[index])).toEqual(
    fullBlocks.map((block, index) => block === inputs[index]),
  );
  // The full pass leaves every input's facts as the cached pass left them:
  // nothing the cached pass skipped had anything left to register.
  expect(sameFacts(inputs.map(keptFacts), cachedFacts)).toBe(true);
  // A block normalized anew carries the same facts either way.
  expect(cachedOutputFacts).toEqual(factValues(fullBlocks));
  // Normalized: exactly the blocks it could not reuse. A new package input
  // starts the story over.
  const counters = countersBeforeEach(full.document);
  const restarted =
    earlier !== undefined &&
    (earlier.document.package.styles !== version.package.styles ||
      earlier.document.package.numbering !== version.package.numbering);
  expect(cache.normalizedBlocks(OP_STORIES.MAIN)).toBe(
    blocksToNormalize(
      inputs,
      tokensBefore,
      counters,
      restarted ? undefined : earlier,
    ),
  );
  return {
    document: cached.document,
    blocks: cachedBlocks,
    tokens: cachedBlocks.map((block) =>
      block.type === "paragraph"
        ? paragraphPropertySnapshotToken(block)
        : undefined,
    ),
    counters,
  };
};

/** Numbering instances the document defines. */
const numIdsOf = (document: Document): number[] =>
  document.package.numbering?.nums.map(({ numId }) => numId) ?? [];

let nextParaId = 0;

const level = (ilvl: number): ListLevel => ({
  ilvl,
  start: 1,
  numFmt: "decimal",
  lvlText: `%${String(ilvl + 1)}.`,
  pPr: {
    indentLeft: 720 * (ilvl + 1),
    indentFirstLine: -360,
    hangingIndent: true,
  },
});

/**
 * Three list instances of one abstract list, the second restarting it and
 * the third continuing the second, at three levels, interleaved with plain
 * paragraphs: long enough that a story holds several checkpoints, and shared
 * enough that two instances advance one counter array.
 */
const sharedListsDocument = async (): Promise<Document> => {
  const content: Paragraph[] = Array.from({ length: 48 }, (_, index) => {
    const paragraph: Paragraph = {
      type: "paragraph",
      paraId: (0x70000000 + index).toString(16).toUpperCase(),
      content: [
        {
          type: "run",
          content: [{ type: "text", text: `Item ${String(index)}` }],
        },
      ],
    };
    if (index % 7 !== 6)
      paragraph.formatting = {
        numPr: {
          kind: "reference",
          numId: [1, 2, 3, 2, 3, 1][Math.floor(index / 8)] ?? 1,
          ilvl: index % 3,
        },
      };
    return paragraph;
  });
  const bytes = await createDocx({
    package: {
      numbering: {
        abstractNums: [
          { abstractNumId: 0, levels: [level(0), level(1), level(2)] },
        ],
        nums: [
          { numId: 1, abstractNumId: 0 },
          {
            numId: 2,
            abstractNumId: 0,
            levelOverrides: [{ ilvl: 0, startOverride: 5 }],
          },
          // Met first after the restarted instance, it continues its counters.
          { numId: 3, abstractNumId: 0 },
        ],
      },
      document: { content },
    },
  });
  return parseDocx(
    (
      await prepareCanonicalDocxInput(new Uint8Array(bytes).slice().buffer)
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
};

/**
 * Lists of one abstract list numbered by direct references and by a
 * paragraph style, a restarting instance among them: a style-numbered
 * paragraph advances the latest instance of the abstract list, so instances
 * trade counter arrays and arrays fall out of use.
 */
const styledListsDocument = async (): Promise<Document> => {
  const content: Paragraph[] = Array.from({ length: 40 }, (_, index) => {
    const paragraph: Paragraph = {
      type: "paragraph",
      paraId: (0x40000000 + index).toString(16).toUpperCase(),
      content: [
        {
          type: "run",
          content: [{ type: "text", text: `Item ${String(index)}` }],
        },
      ],
    };
    if (index % 7 === 6) return paragraph;
    paragraph.formatting =
      index % 3 === 0
        ? { styleId: "ListStyle" }
        : {
            numPr: {
              kind: "reference",
              numId: [2, 3, 1, 4][Math.floor(index / 5) % 4] ?? 1,
              ilvl: index % 2,
            },
          };
    return paragraph;
  });
  const bytes = await createDocx({
    package: {
      styles: {
        styles: [
          {
            styleId: "ListStyle",
            type: "paragraph",
            pPr: { numPr: paragraphNumberingReference({ numId: 4, ilvl: 0 }) },
          },
        ],
      },
      numbering: {
        abstractNums: [
          { abstractNumId: 0, levels: [level(0), level(1), level(2)] },
        ],
        nums: [
          { numId: 1, abstractNumId: 0 },
          {
            numId: 2,
            abstractNumId: 0,
            levelOverrides: [{ ilvl: 0, startOverride: 3 }],
          },
          { numId: 3, abstractNumId: 0 },
          { numId: 4, abstractNumId: 0 },
        ],
      },
      document: { content },
    },
  });
  return parseDocx(
    (
      await prepareCanonicalDocxInput(new Uint8Array(bytes).slice().buffer)
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
};

const generatedParagraphArbitrary = fc.record({
  kind: fc.constantFrom("reference", "reference", "style", "plain", "control"),
  numId: fc.integer({ min: 1, max: 5 }),
  ilvl: fc.nat({ max: 2 }),
});

/**
 * A short story of paragraphs numbered by references to five instances of
 * one abstract list (two of them restarting it), by a paragraph style, or
 * not at all, some numbered paragraphs inside a content control as a parsed
 * package holds them: instances meet, resume and trade counter arrays in
 * every order.
 */
const generatedListsDocument = (
  paragraphs: readonly fc.RecordValue<typeof generatedParagraphArbitrary>[],
): Document => {
  const document = generatedListsModel(paragraphs);
  const numbering = document.package.numbering;
  if (numbering === undefined) return document;
  const map = getCachedNumberingMap(numbering);
  document.package.document.content = document.package.document.content.map(
    (block) => {
      const paragraph = block.type === "blockSdt" ? block.content.at(0) : null;
      const numPr =
        paragraph?.type === "paragraph" && paragraph.formatting?.numPr;
      if (!paragraph || !numPr || numPr.kind !== "reference") return block;
      const rendering = computeListRendering(numPr, map);
      return rendering === null
        ? block
        : { ...block, content: [{ ...paragraph, listRendering: rendering }] };
    },
  );
  return document;
};

const generatedListsModel = (
  paragraphs: readonly fc.RecordValue<typeof generatedParagraphArbitrary>[],
): Document => ({
  package: {
    styles: {
      styles: [
        {
          styleId: "ListStyle",
          type: "paragraph",
          pPr: { numPr: paragraphNumberingReference({ numId: 4, ilvl: 0 }) },
        },
      ],
    },
    numbering: {
      abstractNums: [
        { abstractNumId: 0, levels: [level(0), level(1), level(2)] },
      ],
      nums: [
        { numId: 1, abstractNumId: 0 },
        {
          numId: 2,
          abstractNumId: 0,
          levelOverrides: [{ ilvl: 0, startOverride: 3 }],
        },
        { numId: 3, abstractNumId: 0 },
        { numId: 4, abstractNumId: 0 },
        {
          numId: 5,
          abstractNumId: 0,
          levelOverrides: [{ ilvl: 1, startOverride: 2 }],
        },
      ],
    },
    document: {
      content: paragraphs.map(({ kind, numId, ilvl }, index) => {
        const paragraph: Paragraph = {
          type: "paragraph",
          paraId: (0x30000000 + index).toString(16).toUpperCase(),
          content: [
            {
              type: "run",
              content: [{ type: "text", text: `Item ${String(index)}` }],
            },
          ],
        };
        if (kind === "style") paragraph.formatting = { styleId: "ListStyle" };
        if (kind === "reference" || kind === "control")
          paragraph.formatting = {
            numPr: { kind: "reference", numId, ilvl },
          };
        return kind === "control"
          ? ({
              type: "blockSdt",
              properties: { sdtType: "richText", id: index + 1 },
              content: [paragraph],
            } satisfies BlockContent)
          : paragraph;
      }),
    },
  },
});

const listEditArbitrary = fc.record({
  kind: fc.constantFrom(
    "number",
    "unnumber",
    "level",
    "style",
    "model",
    "again",
    "facts",
    "register",
  ),
  at: fc.nat(),
  numId: fc.nat(),
  ilvl: fc.nat({ max: 2 }),
  model: modelEditArbitrary,
});
type ListEdit = fc.RecordValue<typeof listEditArbitrary>;

/** A new version of the story with one paragraph numbered, renumbered or unnumbered, or one model edit. */
const editLists = (document: Document, edit: ListEdit): Document => {
  if (
    edit.kind === "again" ||
    edit.kind === "facts" ||
    edit.kind === "register"
  )
    return document;
  if (edit.kind === "model") {
    // Inserted paragraphs take an identity, as every paragraph of a session has.
    const edited = editModel(document, edit.model);
    for (const block of mainBlocks(edited))
      if (block.type === "paragraph" && block.paraId === undefined)
        block.paraId = (0x60000000 + nextParaId++).toString(16).toUpperCase();
    return edited;
  }
  const blocks = [...mainBlocks(document)];
  const paragraphs = blocks.flatMap((block, index) =>
    block.type === "paragraph" ? [index] : [],
  );
  const index = paragraphs[edit.at % Math.max(1, paragraphs.length)];
  const block = index === undefined ? undefined : blocks[index];
  if (index === undefined || block?.type !== "paragraph") return document;
  const numIds = numIdsOf(document);
  const numId = numIds[edit.numId % Math.max(1, numIds.length)];
  const current = block.formatting?.numPr;
  let formatting: Paragraph["formatting"];
  switch (edit.kind) {
    case "number":
      if (numId === undefined) return document;
      formatting = {
        ...block.formatting,
        numPr: { kind: "reference", numId, ilvl: edit.ilvl },
      };
      break;
    case "unnumber": {
      formatting = { ...block.formatting };
      delete formatting.numPr;
      break;
    }
    case "level":
      if (current?.kind !== "reference") return document;
      formatting = {
        ...block.formatting,
        numPr: { ...current, ilvl: edit.ilvl },
      };
      break;
    // Numbered by the style instead of a reference, or no longer by it.
    case "style": {
      formatting = { ...block.formatting };
      if (formatting.styleId === "ListStyle") delete formatting.styleId;
      else {
        formatting.styleId = "ListStyle";
        delete formatting.numPr;
      }
      break;
    }
    default:
      return edit.kind satisfies never;
  }
  blocks[index] = { ...block, formatting };
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content: blocks },
    },
  };
};

/** `document` with the main story block `from` replaced by `to`. */
const replacingMainBlock = (
  document: Document,
  from: BlockContent,
  to: BlockContent,
): Document => ({
  ...document,
  package: {
    ...document.package,
    document: {
      ...document.package.document,
      content: mainBlocks(document).map((block) =>
        block === from ? to : block,
      ),
    },
  },
});

describe("a list normalization reusing blocks earlier runs left", () => {
  test("equals a normalization of every block over story versions", async () => {
    const sources = await parsedSources();
    await assertProperty(
      fc.property(
        fc.constantFrom(
          ...sources,
          { id: "shared-lists", document: await sharedListsDocument() },
          { id: "styled-lists", document: await styledListsDocument() },
        ),
        fc.integer({ min: 1, max: 5 }),
        fc.array(listEditArbitrary, { minLength: 1, maxLength: 20 }),
        ({ document }, checkpointSpacing, edits) => {
          const cache = new ListNormalizationCache({ checkpointSpacing });
          let input = structuredClone(document);
          let last = normalizedBothWays(input, cache);
          let version = last.document;
          for (const edit of edits) {
            switch (edit.kind) {
              // The version before normalization once more: blocks the last
              // run changed come back as they were.
              case "again":
                break;
              // The facts kept for a paragraph change: its registration, or its
              // restored projection on a new paragraph in its place.
              case "facts":
              case "register": {
                const paragraphs = mainBlocks(version).filter(
                  (block): block is Paragraph => block.type === "paragraph",
                );
                const paragraph =
                  paragraphs[edit.at % Math.max(1, paragraphs.length)];
                if (paragraph === undefined) break;
                if (edit.kind === "register")
                  registerParagraphPropertySnapshot({
                    paragraph,
                    snapshot: { indentLeft: 1 + edit.ilvl },
                  });
                else {
                  // As history restores a paragraph's facts from another
                  // paragraph's: on a new paragraph in its place once a
                  // normalization kept it.
                  const restored = carryParagraphPropertySource({
                    target: paragraph,
                    source:
                      paragraphs[(edit.at + edit.numId) % paragraphs.length] ??
                      paragraph,
                    indentation: "restore",
                  });
                  if (restored !== paragraph)
                    version = replacingMainBlock(version, paragraph, restored);
                }
                input = version;
                break;
              }
              default:
                input = editLists(version, edit);
            }
            last = normalizedBothWays(input, cache, last);
            version = last.document;
          }
        },
      ),
      {
        id: "list normalization cache equals full normalization over story versions",
        numRuns: 400,
      },
    );
  });

  test("equals a normalization of every block over generated list stories", async () => {
    await assertProperty(
      fc.property(
        fc.array(generatedParagraphArbitrary, { minLength: 2, maxLength: 24 }),
        fc.integer({ min: 1, max: 4 }),
        fc.array(listEditArbitrary, { minLength: 1, maxLength: 16 }),
        (paragraphs, checkpointSpacing, edits) => {
          const cache = new ListNormalizationCache({ checkpointSpacing });
          let last = normalizedBothWays(
            generatedListsDocument(paragraphs),
            cache,
          );
          for (const edit of edits)
            last = normalizedBothWays(
              edit.kind === "again"
                ? last.document
                : editLists(last.document, edit),
              cache,
              last,
            );
        },
      ),
      {
        id: "list normalization cache equals full normalization over generated list stories",
        numRuns: 600,
      },
    );
  });

  test("equals a normalization of every block over a session's documents", async () => {
    const sources = await parsedSources();
    await assertProperty(
      fc.property(
        fc.constantFrom(...sources),
        fc.integer({ min: 1, max: 5 }),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, checkpointSpacing, edits) => {
          const session = createCanonicalSession(
            structuredClone(document),
          ).unwrap();
          const cache = new ListNormalizationCache({ checkpointSpacing });
          let last = normalizedBothWays(session.document, cache);
          let state = EditorState.create({
            schema: session.projection.doc.type.schema,
            doc: session.projection.doc,
          });
          for (const edit of edits) {
            state = applyEdit(session, state, edit);
            last = normalizedBothWays(session.document, cache, last);
          }
        },
      ),
      {
        id: "list normalization cache equals full normalization over session edits",
        numRuns: 200,
      },
    );
  });
});

/** One numbered list of `items` paragraphs. */
const longListDocument = async (items: number): Promise<Document> => {
  const content: Paragraph[] = Array.from({ length: items }, (_, index) => ({
    type: "paragraph",
    paraId: (0x50000000 + index).toString(16).toUpperCase(),
    formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
    content: [
      {
        type: "run",
        content: [{ type: "text", text: `Item ${String(index)}` }],
      },
    ],
  }));
  const bytes = await createDocx({
    package: {
      numbering: {
        abstractNums: [{ abstractNumId: 0, levels: [level(0)] }],
        nums: [{ numId: 1, abstractNumId: 0 }],
      },
      document: { content },
    },
  });
  return parseDocx(
    (
      await prepareCanonicalDocxInput(new Uint8Array(bytes).slice().buffer)
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
};

/**
 * A long list normalized with `cache`, then once more as normalized: the
 * second normalization reuses every block the first left.
 */
const settledLongList = async (
  cache: ListNormalizationCache,
): Promise<Normalized> => {
  const first = normalizedBothWays(await longListDocument(400), cache);
  const settled = normalizedBothWays(first.document, cache, first);
  expect(cache.normalizedBlocks(OP_STORIES.MAIN)).toBe(0);
  return settled;
};

/** `document` with main story block `index` replaced by what `edit` makes of it. */
const editingMainBlock = (
  document: Document,
  index: number,
  edit: (paragraph: Paragraph) => Paragraph,
): Document => {
  const block = mainBlocks(document)[index];
  if (block?.type !== "paragraph") throw new Error("Not a paragraph.");
  return replacingMainBlock(document, block, edit(block));
};

describe("a list normalization after one edit to a long list", () => {
  test("normalizes the edited paragraph alone when no counter changes", async () => {
    const cache = new ListNormalizationCache();
    let last = await settledLongList(cache);
    for (const index of [0, 37, 200, 399]) {
      const edited = editingMainBlock(last.document, index, (paragraph) => ({
        ...paragraph,
        content: [{ type: "run", content: [{ type: "text", text: "Edited" }] }],
      }));
      last = normalizedBothWays(edited, cache, last);
      expect(cache.normalizedBlocks(OP_STORIES.MAIN)).toBe(1);
    }
  });

  test("normalizes the paragraphs after it while their counters change", async () => {
    const cache = new ListNormalizationCache();
    const last = await settledLongList(cache);
    const unnumbered = editingMainBlock(last.document, 100, (paragraph) => {
      const formatting = { ...paragraph.formatting };
      delete formatting.numPr;
      return { ...paragraph, formatting };
    });
    normalizedBothWays(unnumbered, cache, last);
    // The paragraph and every later item, each now one lower.
    expect(cache.normalizedBlocks(OP_STORIES.MAIN)).toBe(300);
  });
});
