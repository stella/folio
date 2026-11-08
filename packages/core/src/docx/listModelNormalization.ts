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
  withBodyContent,
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
  Endnote,
  Footnote,
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
  type ListCounterSnapshot,
  restoreListCounters,
  sameListCounters,
  snapshotListCounters,
} from "./listMarkerComputation";
import { listRenderingDefinitionsMatch } from "./listRendering";
import { computeListRendering, getCachedNumberingMap } from "./numberingParser";
import {
  assignParagraphIndentationProjection,
  cloneParagraphWithPropertySource,
  paragraphFormattingWithAuthoredIndentation,
  paragraphPropertySourceState,
  sameParagraphIndentationProjection,
} from "./paragraphPropertySource";

type CanonicalListNormalization = { document: Document; inverse: DocumentOp[] };

type StyleResolver = ReturnType<typeof createStyleResolver>;

/** The package fields a block's normalization reads; the rest of it is never read. */
type ListNormalizationPackageInputs = Pick<
  Document["package"],
  "styles" | "numbering"
>;

/** One block as an earlier normalization of its story met it. */
type BlockRecord = {
  input: BlockContent;
  /** Normalizing it returned it unchanged and recorded no inverse. */
  settled: boolean;
  /** The paragraph's facts kept outside its fields, as normalizing left them. */
  sourceState: readonly unknown[] | undefined;
  /** The snapshot registration normalizing left for the paragraph. */
  snapshotToken: object | undefined;
};

type StoryRun = {
  records: readonly BlockRecord[];
  /** The counters before the block at an index, kept every few blocks normalized. */
  checkpoints: ReadonlyMap<number, ListCounterSnapshot>;
};

const CHECKPOINT_SPACING = 32;

const sameState = (left: readonly unknown[], right: readonly unknown[]) =>
  left.length === right.length &&
  left.every((value, index) => value === right[index]);

const recordOf = (input: BlockContent, settled: boolean): BlockRecord =>
  input.type === "paragraph"
    ? {
        input,
        settled,
        sourceState: paragraphPropertySourceState(input),
        snapshotToken: paragraphPropertySnapshotToken(input),
      }
    : { input, settled, sourceState: undefined, snapshotToken: undefined };

/** Whether normalizing `block` again would leave it, its facts and its registration as they are. */
const stillSettled = (block: BlockContent, record: BlockRecord): boolean =>
  record.settled &&
  block === record.input &&
  (block.type !== "paragraph" ||
    (record.sourceState !== undefined &&
      sameState(record.sourceState, paragraphPropertySourceState(block)) &&
      paragraphPropertySnapshotToken(block) === record.snapshotToken));

const storyKey = (story: OpStory): string => {
  if (story === OP_STORIES.MAIN) return story;
  return "rId" in story
    ? `${story.kind}:${story.rId}`
    : `${story.kind}:${String(story.id)}`;
};

/**
 * Normalized blocks kept across the normalizations of one session.
 *
 * Normalizing a block reads the block, the package's styles and numbering,
 * and the list counters the blocks before it advanced. A block an edit left
 * the same object, with the same kept facts and snapshot registration, that
 * normalized to itself before, normalizes to itself again from the same
 * counters. So a story is normalized again from the last checkpoint before
 * its first changed block, and only until the counters after its last
 * changed block match those the earlier run had at that place.
 */
export class ListNormalizationCache {
  private inputs: ListNormalizationPackageInputs | undefined;
  private runs = new Map<string, StoryRun>();
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

  /** The blocks of one story normalized, from `counters` as a fresh story starts them. */
  story(
    key: string,
    blocks: readonly BlockContent[],
    counters: ComputeListMarkerOptions,
    normalize: (block: BlockContent) => {
      output: BlockContent;
      settled: boolean;
    },
  ): BlockContent[] {
    const old = this.runs.get(key);
    const oldRecords = old?.records ?? [];
    const oldCheckpoints = old?.checkpoints ?? new Map();
    const length = blocks.length;
    const shift = length - oldRecords.length;
    const settledAt = (index: number, oldIndex: number): boolean => {
      const record = oldRecords[oldIndex];
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
      suffix < oldRecords.length - prefix &&
      settledAt(length - 1 - suffix, oldRecords.length - 1 - suffix)
    )
      suffix += 1;
    const outputs: BlockContent[] = [];
    const records: BlockRecord[] = [];
    const checkpoints = new Map<number, ListCounterSnapshot>();
    const keep = (from: number, to: number, offset: number) => {
      for (let index = from; index < to; index += 1) {
        const record =
          oldRecords[index - offset] ??
          panic("A kept list normalization lost a block.");
        outputs.push(record.input);
        records.push(record);
      }
    };
    // Resume at the last checkpoint at or before the first changed block: the
    // blocks before it normalize as they did.
    let start = 0;
    for (const index of oldCheckpoints.keys())
      if (index <= prefix && index > start) start = index;
    const resumed = oldCheckpoints.get(start);
    if (start > 0 && resumed !== undefined) {
      restoreListCounters(counters, resumed);
      keep(0, start, 0);
      for (const [index, snapshot] of oldCheckpoints)
        if (index <= start) checkpoints.set(index, snapshot);
    } else start = 0;
    for (let index = start; index < length; index += 1) {
      const oldIndex = index - shift;
      const earlier =
        index >= length - suffix ? oldCheckpoints.get(oldIndex) : undefined;
      const here =
        earlier !== undefined || index % this.spacing === 0
          ? snapshotListCounters(counters)
          : undefined;
      if (
        earlier !== undefined &&
        here !== undefined &&
        sameListCounters(earlier, here)
      ) {
        // Every block from here settled before, from these same counters.
        keep(index, length, shift);
        for (const [at, snapshot] of oldCheckpoints)
          if (at >= oldIndex) checkpoints.set(at + shift, snapshot);
        this.runs.set(key, { records, checkpoints });
        return outputs;
      }
      if (here !== undefined && index % this.spacing === 0)
        checkpoints.set(index, here);
      const block = blocks[index] ?? panic("A story lost a block.");
      const { output, settled } = normalize(block);
      outputs.push(output);
      records.push(recordOf(block, settled));
    }
    this.runs.set(key, { records, checkpoints });
    return outputs;
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

/**
 * One block of a story with its list rendering normalized, the counters
 * advanced past it. The block is returned itself when nothing changed.
 */
const normalizeBlock = (
  paragraph: BlockContent,
  { styles, story, recordInverse }: BlockNormalizationInputs,
  counters: ComputeListMarkerOptions,
): BlockContent => {
  const { numbering } = counters;
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
    computeListMarker(paragraph, counters);
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
  computeListMarker(next, counters);
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
  const normalizeBody = <Body extends DocumentBody>(
    body: Body,
    story: OpStory,
  ): Body => {
    const counters: ComputeListMarkerOptions = {
      numbering,
      listCounters: new Map(),
      abstractCounters: new Map(),
      restartedNumIds: new Set(),
      previousList: { abstractNumId: null, fromStyle: false, numId: null },
    };
    const inputs: BlockNormalizationInputs = {
      styles,
      story,
      recordInverse,
    };
    const content =
      cache === undefined
        ? body.content.map((block) => normalizeBlock(block, inputs, counters))
        : cache.story(storyKey(story), body.content, counters, (block) => {
            const before = recorded;
            const output = normalizeBlock(block, inputs, counters);
            return { output, settled: output === block && recorded === before };
          });
    return content.every(
      (paragraph, index) => paragraph === body.content[index],
    )
      ? body
      : { ...body, ...withBodyContent(body, content) };
  };
  const pkg = document.package;
  const body = normalizeBody(pkg.document, OP_STORIES.MAIN);
  const normalizeParts = <Body extends DocumentBody>(
    parts: Map<string, Body> | undefined,
    kind: "header" | "footer",
  ) => {
    if (parts === undefined) return parts;
    const normalized = new Map(
      [...parts].map(([id, part]) => [
        id,
        normalizeBody(part, { kind, rId: id }),
      ]),
    );
    return [...normalized].every(([id, part]) => part === parts.get(id))
      ? parts
      : normalized;
  };
  const normalizeNotes = <Body extends Footnote | Endnote>(
    notes: Body[] | undefined,
    kind: "footnote" | "endnote",
  ) => {
    if (notes === undefined) return notes;
    const normalized = notes.map((note) =>
      normalizeBody(note, { kind, id: note.id }),
    );
    return normalized.every((note, index) => note === notes[index])
      ? notes
      : normalized;
  };
  const headers = normalizeParts(pkg.headers, "header");
  const footers = normalizeParts(pkg.footers, "footer");
  const footnotes = normalizeNotes(pkg.footnotes, "footnote");
  const endnotes = normalizeNotes(pkg.endnotes, "endnote");
  if (
    body === pkg.document &&
    headers === pkg.headers &&
    footers === pkg.footers &&
    footnotes === pkg.footnotes &&
    endnotes === pkg.endnotes
  )
    return { document, inverse };
  return {
    inverse,
    document: {
      ...document,
      package: {
        ...pkg,
        document: body,
        ...(headers === undefined ? {} : { headers }),
        ...(footers === undefined ? {} : { footers }),
        ...(footnotes === undefined ? {} : { footnotes }),
        ...(endnotes === undefined ? {} : { endnotes }),
      },
    },
  };
};
