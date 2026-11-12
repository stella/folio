import { panic } from "better-result";

import { finalizeParagraphFormatting } from "@stll/docx-core/model";
// PARSE-WARNING-EXEMPT: Recomputes derived list rendering after edits; authored numbering and source XML are preserved.
/** Canonical edits retain authored fields; list rendering follows the current numbering. */
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  EMPTY_PROPERTY_SETS,
  OP_STORIES,
  type OpStory,
  paragraphPropertySnapshotToken,
  registerParagraphPropertySnapshot,
  replaceStoryBody,
  withBodyContent,
  withParagraphRenderings,
} from "@stll/docx-core/ops";

import {
  effectiveParagraphNumberingReference,
  paragraphNumberingSourcesFromBoundary,
} from "../prosemirror/numberingAttr";
import { paragraphIndentationFromFormatting } from "../prosemirror/paragraphIndentation";
import { listIndentationProvenancePatch } from "../prosemirror/styles/resolvedStyleAttrs";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import type {
  BlockContent,
  Document,
  DocumentBody,
  Paragraph,
} from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import {
  foldListNumberFields,
  planParagraphListNumberFold,
  unfoldedListNumberContent,
} from "./foldedListNumberFields";
import {
  computeListMarker,
  type ComputeListMarkerOptions,
  counterCellOf,
  type ListCounterAdvance,
  type ListCounterSnapshot,
  type ListCounterStep,
  listCounterStep,
  replayListCounterStep,
  restoreListCounters,
  snapshotListCounters,
} from "./listMarkerComputation";
import { listRenderingDefinitionsMatch } from "./listRendering";
import { computeListRendering, getCachedNumberingMap } from "./numberingParser";
import {
  assignParagraphIndentationProjection,
  cloneParagraphWithPropertySource,
  paragraphFormattingWithAuthoredIndentation,
  sameParagraphIndentationProjection,
  sealParagraphFacts,
} from "./paragraphPropertySource";

type CanonicalListNormalization = { document: Document; inverse: DocumentOp[] };

type StyleResolver = ReturnType<typeof createStyleResolver>;

/** The package fields a block's normalization reads; the rest of it is never read. */
type ListNormalizationPackageInputs = Pick<
  Document["package"],
  "styles" | "numbering"
>;

/** One block as an earlier normalization of its story left it. */
type BlockRecord = {
  input: BlockContent;
  /** Normalizing it returns it unchanged and records no inverse. */
  settled: boolean;
  /** The snapshot registration normalizing left for the paragraph. */
  snapshotToken: object | undefined;
  /**
   * What normalizing it changed in the counters: one step per paragraph it
   * holds, in document order, none for a block holding no counted paragraph.
   */
  steps: readonly ListCounterStep[];
  /** The counters before it, kept every few blocks. */
  before: ListCounterSnapshot | undefined;
};

const CHECKPOINT_SPACING = 32;

/** The record of `input`, whose kept facts are fixed from now on. */
const recordOf = (
  input: BlockContent,
  settled: boolean,
  steps: readonly ListCounterStep[],
  before: ListCounterSnapshot | undefined,
): BlockRecord => {
  if (input.type !== "paragraph")
    return { input, settled, snapshotToken: undefined, steps, before };
  sealParagraphFacts(input);
  return {
    input,
    settled,
    snapshotToken: paragraphPropertySnapshotToken(input),
    steps,
    before,
  };
};

/** Whether normalizing `block` again would leave it, its facts and its registration as they are. */
const stillSettled = (block: BlockContent, record: BlockRecord): boolean =>
  record.settled &&
  block === record.input &&
  (block.type !== "paragraph" ||
    paragraphPropertySnapshotToken(block) === record.snapshotToken);

const storyKey = (story: OpStory): string => {
  if (story === OP_STORIES.MAIN) return story;
  return "rId" in story
    ? `${story.kind}:${story.rId}`
    : `${story.kind}:${String(story.id)}`;
};

const freshCounters = (
  numbering: ComputeListMarkerOptions["numbering"],
): ComputeListMarkerOptions => ({
  numbering,
  listCounters: new Map(),
  abstractCounters: new Map(),
  restartedNumIds: new Set(),
  previousList: { abstractNumId: null, fromStyle: false, numId: null },
});

/**
 * The counters before block `index` of a run, from the nearest kept
 * snapshot at or before it and the steps of the blocks between.
 */
const countersAt = (
  records: readonly BlockRecord[],
  index: number,
  counters: ComputeListMarkerOptions,
): number => {
  let from = Math.min(index, records.length);
  while (from > 0 && records[from]?.before === undefined) from -= 1;
  const before = records[from]?.before;
  if (before !== undefined) restoreListCounters(counters, before);
  else from = 0;
  for (let at = from; at < index; at += 1)
    for (const step of (
      records[at] ?? panic("A list normalization lost a block.")
    ).steps)
      replayListCounterStep(counters, step);
  return from;
};

/**
 * One state of the counters, its arrays by cell. Only arrays an instance
 * or abstract list is bound to are held: an array nothing reaches any more
 * can never be advanced again.
 */
class CounterView {
  private readonly bindings = new Map<number, number>();
  private readonly arrays = new Map<number, readonly number[]>();

  readonly counters: ComputeListMarkerOptions;

  constructor(counters: ComputeListMarkerOptions) {
    this.counters = counters;
    for (const values of counters.listCounters.values()) this.bind(values);
    for (const values of counters.abstractCounters.values()) this.bind(values);
  }

  private bind(values: readonly number[]): void {
    const cell = counterCellOf(values);
    this.bindings.set(cell, (this.bindings.get(cell) ?? 0) + 1);
    this.arrays.set(cell, values);
  }

  private unbind(values: readonly number[]): void {
    const cell = counterCellOf(values);
    const left =
      (this.bindings.get(cell) ??
        panic("A list counter array was not bound.")) - 1;
    if (left > 0) this.bindings.set(cell, left);
    else {
      this.bindings.delete(cell);
      this.arrays.delete(cell);
    }
  }

  /** Follow the bindings a paragraph's step changed. */
  advanced(advance: ListCounterAdvance | null): void {
    if (advance === null) return;
    if (advance.listBefore !== advance.counters) {
      if (advance.listBefore !== undefined) this.unbind(advance.listBefore);
      this.bind(advance.counters);
    }
    if (
      advance.abstractNumId !== null &&
      advance.abstractBefore !== advance.counters
    ) {
      if (advance.abstractBefore !== undefined)
        this.unbind(advance.abstractBefore);
      this.bind(advance.counters);
    }
  }

  cells(): IterableIterator<number> {
    return this.arrays.keys();
  }

  cell(cell: number): readonly number[] | undefined {
    return this.arrays.get(cell);
  }
}

/**
 * One part of the counters: the previous-list state (`p`), the array bound
 * to an instance (`n`) or an abstract list (`a`), whether an instance
 * restarted (`r`), or the values of one array (`c`). Two states equal in
 * every part number every later paragraph alike.
 */
type CounterPart = string;

const cellOrNone = (values: readonly number[] | undefined) =>
  values === undefined ? undefined : counterCellOf(values);

const samePart = (
  part: CounterPart,
  left: CounterView,
  right: CounterView,
): boolean => {
  const l = left.counters;
  const r = right.counters;
  if (part === "p")
    return (
      l.previousList.abstractNumId === r.previousList.abstractNumId &&
      l.previousList.fromStyle === r.previousList.fromStyle &&
      l.previousList.numId === r.previousList.numId
    );
  const id = Number(part.slice(1));
  switch (part[0]) {
    case "n":
      return (
        cellOrNone(l.listCounters.get(id)) ===
        cellOrNone(r.listCounters.get(id))
      );
    case "a":
      return (
        cellOrNone(l.abstractCounters.get(id)) ===
        cellOrNone(r.abstractCounters.get(id))
      );
    case "r":
      return l.restartedNumIds.has(id) === r.restartedNumIds.has(id);
    default: {
      const a = left.cell(id);
      const b = right.cell(id);
      if (a === undefined || b === undefined) return a === b;
      return (
        a.length === b.length &&
        a.every((value, level) => Object.is(value, b[level]))
      );
    }
  }
};

/** The parts a step can have changed. */
const partsOf = (advance: ListCounterAdvance | null): CounterPart[] => {
  if (advance === null) return ["p"];
  const parts: CounterPart[] = [
    "p",
    `n${String(advance.numId)}`,
    `r${String(advance.numId)}`,
    `c${String(counterCellOf(advance.counters))}`,
  ];
  if (advance.abstractNumId !== null)
    parts.push(`a${String(advance.abstractNumId)}`);
  if (advance.listBefore !== undefined)
    parts.push(`c${String(counterCellOf(advance.listBefore))}`);
  if (advance.abstractBefore !== undefined)
    parts.push(`c${String(counterCellOf(advance.abstractBefore))}`);
  return parts;
};

/** Every part in which two states differ. */
const differingParts = (
  left: CounterView,
  right: CounterView,
): Set<CounterPart> => {
  const parts = new Set<CounterPart>(["p"]);
  for (const view of [left, right]) {
    for (const numId of view.counters.listCounters.keys())
      parts.add(`n${String(numId)}`);
    for (const abstractNumId of view.counters.abstractCounters.keys())
      parts.add(`a${String(abstractNumId)}`);
    for (const numId of view.counters.restartedNumIds)
      parts.add(`r${String(numId)}`);
    for (const cell of view.cells()) parts.add(`c${String(cell)}`);
  }
  for (const part of parts) if (samePart(part, left, right)) parts.delete(part);
  return parts;
};

/**
 * Normalized blocks kept across the normalizations of one session.
 *
 * Normalizing a block reads the block, the package's styles and numbering,
 * and the list counters the blocks before it advanced. A block an edit left
 * the same object, with the same kept facts and snapshot registration, that
 * normalized to itself before, normalizes to itself again from the same
 * counters. So a story is normalized again from its first changed block,
 * with the counters the earlier run had there, and after its last changed
 * block only while the counters differ from those the earlier run had at
 * the same block.
 */
export class ListNormalizationCache {
  private inputs: ListNormalizationPackageInputs | undefined;
  private runs = new Map<string, readonly BlockRecord[]>();
  private readonly normalized = new Map<string, number>();
  private readonly spacing: number;

  constructor({
    checkpointSpacing = CHECKPOINT_SPACING,
  }: { checkpointSpacing?: number } = {}) {
    this.spacing = checkpointSpacing;
  }

  /** Start over when the package inputs are not the ones the kept runs read. */
  use(inputs: ListNormalizationPackageInputs): void {
    if (
      this.inputs !== undefined &&
      this.inputs.styles === inputs.styles &&
      this.inputs.numbering === inputs.numbering
    )
      return;
    this.inputs = inputs;
    this.runs = new Map();
  }

  /** How many blocks of `story` its latest normalization normalized. */
  normalizedBlocks(story: OpStory): number {
    return this.normalized.get(storyKey(story)) ?? 0;
  }

  /**
   * The blocks of one story normalized, from `counters` as a fresh story
   * starts them: the place of each block that normalized to another one, and
   * that one.
   */
  story(
    story: OpStory,
    blocks: readonly BlockContent[],
    counters: ComputeListMarkerOptions,
    normalize: (
      block: BlockContent,
      observe: (advance: ListCounterAdvance | null) => void,
    ) => {
      output: BlockContent;
      settled: boolean;
    },
  ): (readonly [number, BlockContent])[] {
    const key = storyKey(story);
    const old = this.runs.get(key) ?? [];
    const length = blocks.length;
    const shift = length - old.length;
    const settledAt = (index: number, oldIndex: number): boolean => {
      const record = old[oldIndex];
      const block = blocks[index];
      return (
        record !== undefined &&
        block !== undefined &&
        stillSettled(block, record)
      );
    };
    let prefix = 0;
    while (prefix < length && settledAt(prefix, prefix)) prefix += 1;
    let suffix = 0;
    while (
      suffix < length - prefix &&
      suffix < old.length - prefix &&
      settledAt(length - 1 - suffix, old.length - 1 - suffix)
    )
      suffix += 1;
    // The blocks before the first changed one stay, and so do the counters
    // the earlier run had there.
    let sinceCheckpoint = prefix - countersAt(old, prefix, counters);
    let records = old.slice(0, prefix);
    const replaced: (readonly [number, BlockContent])[] = [];
    let normalized = 0;
    const next = (index: number): (ListCounterAdvance | null)[] => {
      const block = blocks[index] ?? panic("A story lost a block.");
      const before =
        sinceCheckpoint >= this.spacing
          ? snapshotListCounters(counters)
          : undefined;
      if (before !== undefined) sinceCheckpoint = 0;
      const advances: (ListCounterAdvance | null)[] = [];
      const steps: ListCounterStep[] = [];
      const { output, settled } = normalize(block, (advance) => {
        advances.push(advance);
        steps.push(listCounterStep(counters, advance));
      });
      sinceCheckpoint += 1;
      normalized += 1;
      if (output !== block) replaced.push([index, output]);
      // A block normalized to a new one is kept as that one: normalizing it
      // again from the same counters returns it unchanged.
      records.push(
        recordOf(output, output === block ? settled : true, steps, before),
      );
      return advances;
    };
    let index = prefix;
    for (; index < length - suffix; index += 1) next(index);
    if (index < length) {
      // Every block from here settled before; it settles again from where
      // the counters are those the earlier run met at it.
      const earlier = freshCounters(counters.numbering);
      countersAt(old, index - shift, earlier);
      const live = new CounterView(counters);
      const was = new CounterView(earlier);
      const differing = differingParts(live, was);
      while (differing.size > 0 && index < length) {
        const advances = next(index);
        for (const advance of advances) live.advanced(advance);
        const record =
          old[index - shift] ??
          panic("A kept list normalization lost a block.");
        const advancesBefore = record.steps.map((step) =>
          replayListCounterStep(earlier, step),
        );
        for (const advance of advancesBefore) was.advanced(advance);
        for (const part of [...advances, ...advancesBefore].flatMap(partsOf))
          if (samePart(part, live, was)) differing.delete(part);
          else differing.add(part);
        index += 1;
      }
      records = records.concat(old.slice(index - shift));
    }
    this.runs.set(key, records);
    this.normalized.set(key, normalized);
    return replaced;
  }
}

type ListNormalizationOptions = {
  /**
   * Whether to build the operations that undo the normalization. Only an edit
   * session needs them, and they name paragraphs by identity; a save that only
   * writes the normalized document must work on a package whose paragraphs
   * carry none.
   */
  inverse?: "record" | "omit";
  /**
   * Blocks an earlier normalization of the same session left as they were,
   * reused where the counters before them are the same.
   */
  cache?: ListNormalizationCache;
};

/** What one block's normalization reads besides the block and the counters before it. */
type BlockNormalizationInputs = {
  styles: StyleResolver;
  story: OpStory;
  recordInverse: (operation: () => DocumentOp) => void;
};

/** Called after each paragraph advanced the counters, with what it advanced. */
type ObserveAdvance = (advance: ListCounterAdvance | null) => void;

/**
 * Advance the counters through the paragraphs a content control or custom
 * XML wrapper holds, in document order, as the parser counts them: they
 * continue the story's lists, while a table's cells count their own. Their
 * renderings stay as they are; the counters are computed on copies.
 */
const advanceThroughContainer = (
  blocks: readonly BlockContent[],
  counters: ComputeListMarkerOptions,
  observe: ObserveAdvance,
): void => {
  for (const block of blocks) {
    if (block.type === "blockSdt" || block.type === "blockCustomXml") {
      advanceThroughContainer(block.content, counters, observe);
      continue;
    }
    if (block.type !== "paragraph") continue;
    const rendering = block.listRendering;
    observe(
      computeListMarker(
        rendering === undefined
          ? block
          : cloneParagraphWithPropertySource(block, {
              listRendering: {
                ...rendering,
                marker: rendering.markerTemplate ?? rendering.marker,
              },
            }),
        counters,
      ),
    );
  }
};

/**
 * One block of a story with its list rendering normalized, the counters
 * advanced past it and every paragraph it holds. The block is returned
 * itself when nothing changed.
 */
const normalizeBlock = (
  paragraph: BlockContent,
  { styles, story, recordInverse }: BlockNormalizationInputs,
  counters: ComputeListMarkerOptions,
  observe: ObserveAdvance,
): BlockContent => {
  const { numbering } = counters;
  if (paragraph.type === "blockSdt" || paragraph.type === "blockCustomXml") {
    advanceThroughContainer(paragraph.content, counters, observe);
    return paragraph;
  }
  if (paragraph.type !== "paragraph") return paragraph;
  registerParagraphPropertySnapshot({
    paragraph,
    snapshot: paragraphFormattingWithAuthoredIndentation(paragraph),
  });
  const formatting = paragraph.formatting;
  const next = cloneParagraphWithPropertySource(paragraph, {});
  const styleNumbering = formatting?.styleId
    ? styles.resolveParagraphStyle(formatting.styleId).paragraphFormatting
        ?.numPr
    : undefined;
  if (
    styleNumbering !== undefined &&
    (formatting?.numPr === undefined || formatting.numPr.kind === "levelOnly")
  ) {
    next.formatting = { ...formatting, numPrFromStyle: styleNumbering };
  }
  // A paragraph that owns its reference, or whose named style supplies no
  // numbering, keeps no style-derived numbering from an earlier style. Without
  // a style ID the inherited reference is authored and stays.
  if (
    formatting?.numPrFromStyle !== undefined &&
    ((formatting.styleId !== undefined && styleNumbering === undefined) ||
      formatting.numPr?.kind === "reference" ||
      formatting.numPr?.kind === "none")
  ) {
    next.formatting = { ...formatting };
    delete next.formatting.numPrFromStyle;
  }
  // Derived provenance still belongs to exact inverse preconditions.
  if (
    canonicalJson(next.formatting?.numPrFromStyle) !==
    canonicalJson(formatting?.numPrFromStyle)
  ) {
    recordInverse(() => ({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story,
      blockId:
        paragraph.paraId ??
        panic("List normalization requires paragraph identities."),
      patch: { numPrFromStyle: formatting?.numPrFromStyle ?? null },
      whenEmpty:
        formatting === undefined
          ? EMPTY_PROPERTY_SETS.OMIT
          : EMPTY_PROPERTY_SETS.KEEP,
    }));
  }
  const numPr = effectiveParagraphNumberingReference(
    paragraphNumberingSourcesFromBoundary({
      numPr: next.formatting?.numPr,
      numPrFromStyle: next.formatting?.numPrFromStyle,
    }),
  );
  if (
    paragraph.listRendering === undefined &&
    formatting?.numPrFromStyle === undefined &&
    next.formatting?.numPrFromStyle === undefined &&
    (numPr?.kind !== "reference" || numbering === null)
  ) {
    observe(computeListMarker(paragraph, counters));
    return paragraph;
  }
  const definition =
    numPr?.kind === "reference" && numbering !== null
      ? computeListRendering(numPr, numbering)
      : null;
  const cached = paragraph.listRendering;
  const rendering =
    definition !== null &&
    cached !== undefined &&
    listRenderingDefinitionsMatch(cached, definition)
      ? { ...cached }
      : definition;
  const authoredFormatting = paragraphFormattingWithAuthoredIndentation(next);
  const authored = paragraphIndentationFromFormatting(authoredFormatting);
  const inheritance = listIndentationProvenancePatch({
    direct: authored,
    styleFormatting: styles.resolveParagraphStyle(next.formatting?.styleId)
      .paragraphFormatting,
    numberingSource:
      next.formatting?.numPr?.kind === "reference" ||
      next.formatting?.numPrFromStyle === undefined
        ? "paragraph"
        : "style",
    numPr:
      numPr?.kind === "reference"
        ? { numId: numPr.numId, ilvl: numPr.ilvl ?? 0 }
        : undefined,
    numbering,
  });
  assignParagraphIndentationProjection({
    paragraph: next,
    authored: authoredFormatting,
    inherited:
      paragraphIndentationFromFormatting(inheritance._resolvedFormatting) ?? {},
  });
  // Structural inverse preconditions capture the fields before rendering
  // normalization; undo its effective indentation before those inverses.
  if (
    canonicalJson(paragraphIndentationFromFormatting(next.formatting)) !==
    canonicalJson(paragraphIndentationFromFormatting(formatting))
  ) {
    recordInverse(() => ({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story,
      blockId:
        paragraph.paraId ??
        panic(
          "Canonical numbering normalization requires paragraph identities.",
        ),
      patch: {
        indentLeft: formatting?.indentLeft ?? null,
        indentRight: formatting?.indentRight ?? null,
        indentFirstLine: formatting?.indentFirstLine ?? null,
        hangingIndent: formatting?.hangingIndent ?? null,
      },
      whenEmpty:
        formatting === undefined
          ? EMPTY_PROPERTY_SETS.OMIT
          : EMPTY_PROPERTY_SETS.KEEP,
    }));
  }
  if (rendering === null) delete next.listRendering;
  else {
    const fields = foldListNumberFields(
      paragraph.content.map(unfoldedListNumberContent),
      () => undefined,
    ).fieldCount;
    if (fields === 0) delete rendering.implicitChildLevelAdvances;
    else rendering.implicitChildLevelAdvances = fields;
    // Counters consume templates, never a marker substituted on an earlier pass.
    const template = rendering.markerTemplate ?? rendering.marker;
    const fold = planParagraphListNumberFold(
      paragraph.content,
      !rendering.isBullet && template !== "" && !template.includes("\t"),
    );
    if (fold.suffix === undefined) delete rendering.foldedMarkerSuffix;
    else rendering.foldedMarkerSuffix = fold.suffix;
    const child = numbering?.getLevel(
      rendering.numId,
      rendering.level + 1,
    )?.pPr;
    if (
      fold.suffix !== undefined &&
      child?.hangingIndent === true &&
      child.indentFirstLine !== undefined &&
      child.indentFirstLine < 0
    )
      rendering.markerSecondSlotOffsetTwips = -child.indentFirstLine;
    else delete rendering.markerSecondSlotOffsetTwips;
    rendering.marker =
      rendering.foldedMarkerSuffix !== undefined && !template.includes("\t")
        ? `${template}\t${rendering.foldedMarkerSuffix}`
        : template;
    next.listRendering = rendering;
  }
  observe(computeListMarker(next, counters));
  registerParagraphPropertySnapshot({
    paragraph: next,
    snapshot: authoredFormatting,
  });
  const finalizedFormatting = finalizeParagraphFormatting(next.formatting);
  if (finalizedFormatting === undefined) delete next.formatting;
  if (
    finalizeParagraphFormatting(formatting) === formatting &&
    sameParagraphIndentationProjection({ left: next, right: paragraph }) &&
    canonicalJson(next.formatting) === canonicalJson(formatting) &&
    canonicalJson(next.listRendering) === canonicalJson(cached)
  ) {
    return paragraph;
  }
  return next;
};

export const normalizeListModelRendering = (
  document: Document,
  { inverse: inverseMode = "record", cache }: ListNormalizationOptions = {},
): CanonicalListNormalization => {
  const inverse: DocumentOp[] = [];
  let recorded = 0;
  const recordInverse = (operation: () => DocumentOp): void => {
    recorded += 1;
    if (inverseMode === "record") inverse.unshift(operation());
  };
  // Everything a block's normalization reads from the package.
  const packageInputs: ListNormalizationPackageInputs = {
    styles: document.package.styles,
    numbering: document.package.numbering,
  };
  cache?.use(packageInputs);
  const styles = createStyleResolver(packageInputs.styles);
  const numbering =
    packageInputs.numbering === undefined
      ? null
      : getCachedNumberingMap(packageInputs.numbering);
  /** The place of each block of one story that normalized to another one, and that one. */
  const replacedIn = (
    body: DocumentBody,
    story: OpStory,
  ): (readonly [number, BlockContent])[] => {
    const counters = freshCounters(numbering);
    const inputs: BlockNormalizationInputs = {
      styles,
      story,
      recordInverse,
    };
    return cache === undefined
      ? body.content.flatMap((block, index) => {
          const output = normalizeBlock(block, inputs, counters, () => {});
          return output === block ? [] : [[index, output] as const];
        })
      : cache.story(story, body.content, counters, (block, observe) => {
          const before = recorded;
          const output = normalizeBlock(block, inputs, counters, observe);
          return { output, settled: output === block && recorded === before };
        });
  };
  /**
   * `document` with one story's replaced blocks in their places. A paragraph
   * normalized to another takes its place through the renderings helper,
   * which keeps the document's standing against the operations contract; a
   * story that holds one paragraph object at two places has its blocks
   * replaced by place instead, and the result is checked anew.
   */
  const withReplaced = (
    target: Document,
    story: OpStory,
    body: DocumentBody,
    replaced: readonly (readonly [number, BlockContent])[],
  ): Document => {
    const renderings = new Map<Paragraph, Paragraph>();
    let byPlace = false;
    for (const [index, output] of replaced) {
      const block = body.content[index];
      if (block?.type !== "paragraph" || output.type !== "paragraph")
        return panic("List normalization replaces only paragraphs.");
      const known = renderings.get(block);
      if (known !== undefined && known !== output) byPlace = true;
      renderings.set(block, output);
    }
    const rendered = byPlace
      ? undefined
      : withParagraphRenderings({ document: target, story, renderings });
    if (rendered !== undefined) return rendered;
    const content = [...body.content];
    for (const [index, output] of replaced) content[index] = output;
    return replaceStoryBody({
      document: target,
      story,
      body: { ...body, ...withBodyContent(body, content) },
    });
  };
  const pkg = document.package;
  // Stories in the order every normalization takes them: the counters and
  // the inverse operations follow it.
  const stories: { story: OpStory; body: DocumentBody }[] = [
    { story: OP_STORIES.MAIN, body: pkg.document },
    ...[...(pkg.headers ?? [])].map(([rId, body]) => ({
      story: { kind: "header" as const, rId },
      body,
    })),
    ...[...(pkg.footers ?? [])].map(([rId, body]) => ({
      story: { kind: "footer" as const, rId },
      body,
    })),
    ...(pkg.footnotes ?? []).map((body) => ({
      story: { kind: "footnote" as const, id: body.id },
      body,
    })),
    ...(pkg.endnotes ?? []).map((body) => ({
      story: { kind: "endnote" as const, id: body.id },
      body,
    })),
  ];
  // Each story's renderings replace its paragraphs in place: a document known
  // to meet the operations contract still does.
  let normalized = document;
  for (const { story, body } of stories)
    normalized = withReplaced(normalized, story, body, replacedIn(body, story));
  return { document: normalized, inverse };
};
