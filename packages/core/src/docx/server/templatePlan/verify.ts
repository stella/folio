/**
 * Verification of an applied template plan. The output package is read back
 * with the same reader that analysed the source and compared, block by block,
 * against what the plan resolves to from the source alone: text, stated
 * emphasis, notes, references, paragraph style and list membership, and the
 * integrity of every id the body uses. A kept template block must read as
 * the template has it.
 *
 * Limits. These are not compared, so a change to them passes:
 * - run formatting other than emphasis and highlight (fonts, sizes, colours)
 *   and paragraph formatting other than the style and the list, on placed
 *   and on kept blocks alike;
 * - the content of headers, footers and text boxes, beyond the fields they
 *   carry;
 * - cell spans and merges of a table;
 * - levels of the numbering part the plan neither patches nor restarts;
 * - what a field shows once it updates: a cached result is compared, the
 *   label a list will render is only predicted.
 */
import { Result } from "better-result";

import { isNumberingReference } from "../../numberingReference";
import { parseSettings } from "../../settingsParser";
import type { XmlElement } from "../../xmlParser";
import { validateDocxConformance } from "../validateDocxConformance";
import {
  type AnalyzedBlock,
  type AnalyzedDocument,
  type AnalyzedNote,
  type AnalyzedNotes,
  analyzeDocument,
  type AnalyzedStory,
} from "./blocks";
import { fieldBookmarkOf } from "./builders";
import {
  linkAddress,
  linkAddressKey,
  type TemplateCharacterStyles,
  templateCharacterStyles,
} from "./emit";
import {
  emphasisSignature,
  readDirectEmphasis,
  signatureHighlight,
  signatureWithoutHighlight,
} from "./emphasis";
import { partFields, readHeaderFooterParts } from "./inspect";
import { styleNumbering } from "./labels";
import { guarded, readPartText } from "./package";
import {
  type ExpectedPiece,
  type ExpectedSegment,
  hyperlinkFieldLink,
  paragraphPieces,
  piecesContent,
  type ResolvedParagraph,
  type ResolvedPlan,
  resolvePlan,
  segmentPieces,
} from "./plan";
import {
  objectFingerprint,
  objectTargets,
  readStoryResources,
  relationshipIdsOf,
  type StoryOrigin,
  type StoryResources,
} from "./resources";
import {
  type AtomKind,
  type AtomLink,
  createStoryState,
  endStory,
  type ParagraphRead,
  readParagraph,
  type StoryField,
  walkBlocks,
} from "./story";
import type {
  ApplyTemplatePlanOptions,
  TemplatePlan,
  TemplatePlanBlockResult,
  TemplatePlanError,
  TemplatePlanFieldLeftover,
  TemplatePlanIssue,
  TemplatePlanNoteCounts,
  TemplatePlanReferenceResult,
  TemplatePlanReport,
  TemplatePlanSpanCounts,
  VerifyTemplatePlanOutputArgs,
} from "./types";
import {
  integerOrNull,
  isW,
  type RawPart,
  readRawPart,
  twipsOrNull,
  walkElements,
  wAttr,
  wChild,
  wChildren,
  wChildVal,
  wName,
} from "./xml";

const PROMPTING_FIELD_KINDS: ReadonlySet<string> = new Set(["FILLIN", "ASK"]);
const NOTE_REFERENCE_KINDS: ReadonlySet<string> = new Set([
  "footnoteReference",
  "endnoteReference",
]);
const BOOKMARK_FIELD_KINDS: ReadonlySet<string> = new Set([
  "REF",
  "PAGEREF",
  "NOTEREF",
]);

/**
 * What a reader gets from one paragraph or note: its text, and per unit of
 * text the emphasis stated and the address linked to; plus the things that
 * take no text, page breaks and objects, in order and with their offsets.
 */
type Content = {
  text: string;
  signatures: string[];
  links: string[];
  marks: string[];
};

const emptyContent = (): Content => ({
  text: "",
  signatures: [],
  links: [],
  marks: [],
});

type ContentUnit = {
  kind: AtomKind;
  text: string;
  signature: string;
  link: string;
  element: XmlElement | null;
  origin: StoryOrigin;
};

const appendUnit = (
  content: Content,
  { kind, text, signature, link, element, origin }: ContentUnit,
  resources: StoryResources,
): void => {
  content.text += text;
  for (let unit = 0; unit < text.length; unit += 1) {
    content.signatures.push(signature);
    content.links.push(link);
  }
  // A mark carries the offset it sits at: a break after the text is not the
  // break before it.
  const at = content.text.length - text.length;
  if (kind === "pageBreak") {
    content.marks.push(`${at}:break:${wAttr(element, "type") ?? "page"}`);
  } else if (kind === "object" && element !== null) {
    content.marks.push(
      `${at}:object:${objectFingerprint(element)}:${objectTargets(element, origin, resources).join(",")}`,
    );
  }
};

/** A paragraph break inside a note. */
const appendParagraphBreak = (content: Content): void => {
  content.text += "\n";
  content.signatures.push("");
  content.links.push("");
};

/** What the source's pieces say the output must hold. */
const expectedContent = (
  pieces: readonly ExpectedPiece[],
  resources: StoryResources,
  content: Content = emptyContent(),
): Content => {
  for (const piece of pieces) {
    appendUnit(
      content,
      {
        kind: piece.kind,
        text: piece.text,
        signature: piece.signature,
        link: linkAddressKey(
          linkAddress(piece.link, resources.relationships[piece.origin]),
        ),
        element: piece.atom?.element ?? null,
        origin: piece.origin,
      },
      resources,
    );
  }
  return content;
};

type OutputContentOptions = {
  read: ParagraphRead;
  origin: StoryOrigin;
  resources: StoryResources;
  styles: TemplateCharacterStyles;
  content?: Content;
};

/**
 * What a paragraph of the output holds, read off its runs. A character style
 * counts as stated emphasis unless it is the one the writer gives a link or
 * a note reference mark, so a style slipped onto a run does not pass unseen.
 */
const outputContent = ({
  read,
  origin,
  resources,
  styles,
  content = emptyContent(),
}: OutputContentOptions): Content => {
  for (const atom of read.atoms) {
    const link = linkAddressKey(
      linkAddress(atom.link, resources.relationships[origin]),
    );
    const isNoteReference = NOTE_REFERENCE_KINDS.has(atom.kind);
    const styleId = wChildVal(atom.runProperties, "rStyle");
    let expectedStyle: string | null = null;
    if (atom.kind === "footnoteReference") {
      expectedStyle = styles.footnoteReference;
    } else if (atom.kind === "endnoteReference") {
      expectedStyle = styles.endnoteReference;
    } else if (link !== "") {
      expectedStyle = styles.hyperlink;
    }
    const stated = isNoteReference
      ? ""
      : emphasisSignature(readDirectEmphasis(atom.runProperties));
    appendUnit(
      content,
      {
        kind: atom.kind,
        text: atom.text,
        signature:
          styleId === null || styleId === expectedStyle
            ? stated
            : `${stated};rStyle=${styleId}`,
        link,
        element: atom.element,
        origin,
      },
      resources,
    );
  }
  return content;
};

const sameList = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length &&
  left.every((entry, index) => entry === right[index]);

/** Whether two contents read the same, emphasis aside. */
const sameReading = (left: Content, right: Content): boolean =>
  left.text === right.text &&
  sameList(left.links, right.links) &&
  sameList(left.marks, right.marks);

/** Maximal runs of units for which `project` gives one non-empty value. */
const countSpans = (
  signatures: readonly string[],
  project: (signature: string) => string,
): number => {
  let count = 0;
  let previous = "";
  for (const signature of signatures) {
    const value = project(signature);
    if (value !== "" && value !== previous) {
      count += 1;
    }
    previous = value;
  }
  return count;
};

const emphasisSpans = (signatures: readonly string[]): number =>
  countSpans(signatures, signatureWithoutHighlight);

const highlightSpans = (signatures: readonly string[]): number =>
  countSpans(signatures, (signature) => signatureHighlight(signature) ?? "");

type SpanTally = { emphasis: number; highlight: number };

const tally = (target: SpanTally, signatures: readonly string[]): void => {
  target.emphasis += emphasisSpans(signatures);
  target.highlight += highlightSpans(signatures);
};

const nestedParagraphs = (
  story: AnalyzedStory,
  topLevelId: string,
): AnalyzedBlock[] =>
  story.order.filter(
    (block) => block.topLevelId === topLevelId && block.read !== null,
  );

const noteReferenceIds = (
  atoms: readonly { kind: string; element: XmlElement | null }[],
  kind: "footnoteReference" | "endnoteReference",
): string[] =>
  atoms.flatMap((atom) => {
    const id = atom.kind === kind ? wAttr(atom.element, "id") : null;
    return id === null ? [] : [id];
  });

type SourceNoteOptions = {
  note: AnalyzedNote;
  source: AnalyzedDocument;
  origin: StoryOrigin;
  fields: readonly StoryField[];
  resources: StoryResources;
};

const sourceNoteContent = ({
  note,
  source,
  origin,
  fields,
  resources,
}: SourceNoteOptions): Content => {
  const content = emptyContent();
  for (const [index, read] of note.paragraphs.entries()) {
    if (index > 0) {
      appendParagraphBreak(content);
    }
    expectedContent(
      paragraphPieces({ read, context: source, origin, fields }),
      resources,
      content,
    );
  }
  return content;
};

type OutputNoteOptions = Omit<OutputContentOptions, "read" | "content"> & {
  note: AnalyzedNote;
};

const outputNoteContent = ({
  note,
  ...options
}: OutputNoteOptions): Content => {
  const content = emptyContent();
  for (const [index, read] of note.paragraphs.entries()) {
    if (index > 0) {
      appendParagraphBreak(content);
    }
    outputContent({ ...options, read, content });
  }
  return content;
};

const countByKind = (
  part: string,
  fields: readonly StoryField[],
): TemplatePlanFieldLeftover[] => {
  const counts = new Map<string, number>();
  for (const { kind } of fields) {
    counts.set(kind, (counts.get(kind) ?? 0) + 1);
  }
  return [...counts].map(([kind, count]) => ({ part, kind, count }));
};

const partText = (part: RawPart): string => {
  const state = createStoryState();
  const texts: string[] = [];
  const readUnder = (container: XmlElement): void => {
    walkBlocks(container, {
      paragraph: (paragraph) => {
        texts.push(readParagraph(paragraph, state).text);
      },
      tableStart: () => {},
      cellStart: () => {},
      tableEnd: () => {},
      unread: () => {},
      tracked: () => {},
    });
  };
  if (isW(part.root, "footnotes") || isW(part.root, "endnotes")) {
    for (const note of part.root.elements ?? []) {
      if (note.type === "element") {
        readUnder(note);
        endStory(state);
      }
    }
  } else {
    readUnder(part.root);
  }
  return texts.join("\n");
};

const occurrences = (text: string, needle: string): number => {
  if (needle === "") {
    return 0;
  }
  let count = 0;
  let from = text.indexOf(needle);
  while (from !== -1) {
    count += 1;
    from = text.indexOf(needle, from + needle.length);
  }
  return count;
};

const CONFORMANCE_RANK = {
  conformant: 2,
  indeterminate: 1,
  invalid: 0,
} as const;

const NUMBER_SWITCH = /\\[nrw]\b/u;

/** The style and the list a paragraph states for itself. */
type StatedProperties = {
  styleId: string | null;
  numId: number | null;
  ilvl: number | null;
};

const statedProperties = (paragraph: XmlElement): StatedProperties => {
  const properties = wChild(paragraph, "pPr");
  const numbering = wChild(properties, "numPr");
  return {
    styleId: wChildVal(properties, "pStyle"),
    numId: integerOrNull(wChildVal(numbering, "numId")),
    ilvl: integerOrNull(wChildVal(numbering, "ilvl")),
  };
};

/** Whether a bookmark of `read` holds all of the paragraph's content. */
const bookmarkEncloses = (
  read: ParagraphRead | null,
  name: string,
): boolean => {
  const start = read?.bookmarks.find((bookmark) => bookmark.name === name);
  const end =
    start === undefined
      ? undefined
      : read?.bookmarkEnds.find(({ id }) => id === start.id);
  return (
    read !== null &&
    start !== undefined &&
    end !== undefined &&
    start.atomIndex === 0 &&
    end.atomIndex === read.atoms.length
  );
};

const segmentLength = (segment: ExpectedSegment): number =>
  piecesContent(segment.kind === "piece" ? [segment.piece] : segment.pieces)
    .text.length;

/** The rows and cells of a table, without the ids of what they hold. */
const tableShape = (block: AnalyzedBlock): string =>
  JSON.stringify(
    (block.block.cells ?? []).map(({ row, col, blockIds }) => [
      row,
      col,
      blockIds.length,
    ]),
  );

type ReportArgs = {
  source: AnalyzedDocument;
  template: AnalyzedDocument;
  output: Uint8Array;
  plan: TemplatePlan;
  resolved: ResolvedPlan;
  options: ApplyTemplatePlanOptions | undefined;
};

/** Read `output` back and report how it compares with the resolved plan. */
export const reportTemplatePlanOutput = async ({
  source,
  template,
  output,
  plan,
  resolved,
  options,
}: ReportArgs): Promise<Result<TemplatePlanReport, TemplatePlanError>> => {
  const analyzedOutput = await analyzeDocument(output);
  if (Result.isError(analyzedOutput)) {
    return analyzedOutput;
  }
  const out = analyzedOutput.value;
  const issues: TemplatePlanIssue[] = [];
  const warnings: string[] = [
    ...source.manifest.warnings.map((warning) => `Source: ${warning}`),
    ...out.story.warnings,
  ];
  const sourceResourcesRead = await readStoryResources(source);
  if (Result.isError(sourceResourcesRead)) {
    return sourceResourcesRead;
  }
  const outputResourcesRead = await readStoryResources(out);
  if (Result.isError(outputResourcesRead)) {
    return outputResourcesRead;
  }
  const templateResourcesRead = await readStoryResources(template);
  if (Result.isError(templateResourcesRead)) {
    return templateResourcesRead;
  }
  const sourceResources = sourceResourcesRead.value;
  const outputResources = outputResourcesRead.value;
  const templateResources = templateResourcesRead.value;
  const characterStyles = templateCharacterStyles(template);

  // The style and the list the plan gives each paragraph. A slot keeps what
  // its template paragraph states; any other paragraph gets what the plan
  // resolves to, with a list its style binds cancelled when the plan says so.
  const slotBlocks = new Map<ResolvedParagraph, AnalyzedBlock>();
  for (const item of resolved.items) {
    if (item.kind === "slot") {
      slotBlocks.set(item.paragraph, item.templateBlock);
    }
  }
  for (const slot of resolved.nestedSlots) {
    slotBlocks.set(slot.paragraph, slot.templateBlock);
  }
  const slottedTemplateBlocks = new Set(slotBlocks.values());
  const plannedProperties = (
    paragraph: ResolvedParagraph,
  ): StatedProperties => {
    const slot = slotBlocks.get(paragraph);
    if (slot !== undefined) {
      return statedProperties(slot.element);
    }
    const { styleId, numbering } = paragraph;
    if (numbering.kind === "direct") {
      return { styleId, numId: numbering.numId, ilvl: numbering.ilvl };
    }
    const cancelsList =
      numbering.kind === "none" &&
      styleNumbering(
        styleId ?? template.defaultStyleId,
        template.styles,
        template.numbering,
      )?.kind === "reference";
    return { styleId, numId: cancelsList ? 0 : null, ilvl: null };
  };
  const checkProperties = (
    planned: StatedProperties,
    stated: StatedProperties,
    blockId: string,
  ): void => {
    if (stated.styleId !== planned.styleId) {
      issues.push({
        code: "style-mismatch",
        message: `Block ${blockId} has the style ${stated.styleId ?? "(none)"}, the plan gives it ${planned.styleId ?? "(none)"}.`,
        blockId,
      });
    }
    if (stated.numId !== planned.numId || stated.ilvl !== planned.ilvl) {
      issues.push({
        code: "numbering-mismatch",
        message: `Block ${blockId} is not in the list, or at the level, the plan gives it.`,
        blockId,
      });
    }
  };

  // Which output block each resolved paragraph was written as.
  const outputOf = new Map<ResolvedParagraph, AnalyzedBlock>();
  const templateOutputIds = new Set<string>();
  const missing = (message: string): void => {
    issues.push({ code: "block-missing", message });
  };
  const assign = (
    paragraph: ResolvedParagraph | undefined,
    block: AnalyzedBlock | undefined,
  ): void => {
    if (paragraph === undefined) {
      return;
    }
    if (block === undefined || block.read === null) {
      missing(
        `Source block ${paragraph.parts.at(0)?.block.block.id ?? "?"} has no paragraph in the output.`,
      );
      return;
    }
    outputOf.set(paragraph, block);
  };
  // A source table and its output: the same rows and cells, and in each cell
  // the same blocks in order.
  const tableOutputs = new Map<string, string>();
  const mapTable = (
    cells: ReadonlyMap<XmlElement, ResolvedParagraph>,
    sourceTable: AnalyzedBlock,
    outputTable: AnalyzedBlock,
  ): void => {
    const tableId = sourceTable.block.id;
    tableOutputs.set(tableId, outputTable.block.id);
    const sourceCells = sourceTable.block.cells ?? [];
    const outputCells = outputTable.block.cells ?? [];
    if (
      sourceCells.length !== outputCells.length ||
      sourceCells.some(
        ({ row, col }, index) =>
          outputCells[index]?.row !== row || outputCells[index]?.col !== col,
      )
    ) {
      missing(`Source table ${tableId} does not keep its rows and cells.`);
      return;
    }
    for (const [index, sourceCell] of sourceCells.entries()) {
      const sourceBlocks = sourceCell.blockIds.map((id) =>
        source.story.byId.get(id),
      );
      const outputBlocks = (outputCells[index]?.blockIds ?? []).map((id) =>
        out.story.byId.get(id),
      );
      // A cell ends in a paragraph; the writer adds an empty one to a cell
      // of the source that did not.
      const last = sourceBlocks.at(-1);
      const closing = last === undefined || last.read === null ? 1 : 0;
      const added = closing === 1 ? outputBlocks.at(-1) : undefined;
      if (
        outputBlocks.length !== sourceBlocks.length + closing ||
        (closing === 1 &&
          (added?.read === undefined ||
            added.read === null ||
            added.read.atoms.length > 0))
      ) {
        missing(`A cell of source table ${tableId} does not keep its blocks.`);
        continue;
      }
      for (const [position, sourceBlock] of sourceBlocks.entries()) {
        const outputBlock = outputBlocks[position];
        if (sourceBlock === undefined) {
          continue;
        }
        if (sourceBlock.read !== null) {
          assign(cells.get(sourceBlock.element), outputBlock);
        } else if (outputBlock === undefined || outputBlock.read !== null) {
          missing(`Source table ${sourceBlock.block.id} is not in the output.`);
        } else {
          mapTable(cells, sourceBlock, outputBlock);
        }
      }
    }
  };
  const outputTopLevel = out.story.topLevel;
  for (const [position, item] of resolved.items.entries()) {
    const block = outputTopLevel[position];
    switch (item.kind) {
      case "template": {
        if (block === undefined || block.block.kind !== item.block.block.kind) {
          missing(
            `Template block ${item.block.block.id} is not where the plan puts it.`,
          );
          break;
        }
        templateOutputIds.add(block.block.id);
        const templateNested = nestedParagraphs(
          template.story,
          item.block.block.id,
        );
        const outputNested = nestedParagraphs(out.story, block.block.id);
        // A kept block reads as the template has it: the same paragraphs,
        // each with its text, links, objects, emphasis, style and list. A
        // frozen field keeps its result, so it reads the same too.
        const keptAsWritten =
          templateNested.length === outputNested.length &&
          tableShape(item.block) === tableShape(block) &&
          templateNested.every((templateParagraph, index) => {
            const outputParagraph = outputNested[index];
            if (
              outputParagraph === undefined ||
              templateParagraph.read === null ||
              outputParagraph.read === null
            ) {
              return false;
            }
            if (slottedTemplateBlocks.has(templateParagraph)) {
              return true;
            }
            const before = statedProperties(templateParagraph.element);
            const after = statedProperties(outputParagraph.element);
            const was = outputContent({
              read: templateParagraph.read,
              origin: "document",
              resources: templateResources,
              styles: characterStyles,
            });
            const is = outputContent({
              read: outputParagraph.read,
              origin: "document",
              resources: outputResources,
              styles: characterStyles,
            });
            return (
              before.styleId === after.styleId &&
              before.numId === after.numId &&
              before.ilvl === after.ilvl &&
              sameReading(was, is) &&
              sameList(was.signatures, is.signatures)
            );
          });
        if (!keptAsWritten) {
          issues.push({
            code: "template-mismatch",
            message: `Template block ${item.block.block.id} does not read as the template has it.`,
            blockId: item.block.block.id,
          });
        }
        for (const slot of resolved.nestedSlots) {
          if (slot.templateBlock.topLevelId !== item.block.block.id) {
            continue;
          }
          assign(
            slot.paragraph,
            outputNested[templateNested.indexOf(slot.templateBlock)],
          );
        }
        break;
      }
      case "paragraph":
        assign(item.paragraph, block);
        break;
      case "slot":
        assign(item.paragraph, block);
        break;
      case "table": {
        if (block === undefined || block.read !== null) {
          missing(`Source table ${item.block.block.id} is not in the output.`);
          break;
        }
        mapTable(item.cells, item.block, block);
        const tableStyle = wChildVal(
          wChild(block.element, "tblPr"),
          "tblStyle",
        );
        if (tableStyle !== item.styleId) {
          issues.push({
            code: "style-mismatch",
            message: `Table ${item.block.block.id} has the style ${tableStyle ?? "(none)"}, the plan gives it ${item.styleId ?? "(none)"}.`,
            blockId: item.block.block.id,
          });
        }
        break;
      }
      default:
        item satisfies never;
    }
  }
  if (outputTopLevel.length !== resolved.items.length) {
    missing(
      `The plan places ${resolved.items.length} block(s), the output holds ${outputTopLevel.length}.`,
    );
  }

  // Text, emphasis and objects of every placed paragraph.
  type Comparison = {
    text: "identical" | "mismatch";
    emphasis: "identical" | "mismatch";
    output: AnalyzedBlock | null;
  };
  const comparisons = new Map<ResolvedParagraph, Comparison>();
  const spans = {
    source: { emphasis: 0, highlight: 0 },
    expected: { emphasis: 0, highlight: 0 },
    output: { emphasis: 0, highlight: 0 },
  } satisfies Record<keyof TemplatePlanSpanCounts, SpanTally>;
  const resolvedParagraphs = [...new Set(resolved.byBlockId.values())];
  for (const paragraph of resolvedParagraphs) {
    tally(spans.expected, paragraph.signatures);
    const block = outputOf.get(paragraph);
    const blockId = paragraph.parts.at(0)?.block.block.id ?? "";
    if (block === undefined || block.read === null) {
      comparisons.set(paragraph, {
        text: "mismatch",
        emphasis: "mismatch",
        output: null,
      });
      continue;
    }
    const expected = expectedContent(
      segmentPieces(paragraph.segments),
      sourceResources,
    );
    const actual = outputContent({
      read: block.read,
      origin: "document",
      resources: outputResources,
      styles: characterStyles,
    });
    tally(spans.output, actual.signatures);
    checkProperties(
      plannedProperties(paragraph),
      statedProperties(block.element),
      blockId,
    );
    const text = sameReading(actual, expected) ? "identical" : "mismatch";
    const emphasis = sameList(actual.signatures, expected.signatures)
      ? "identical"
      : "mismatch";
    // Source text inside a field of the template would go when it updates.
    if (
      block.read.openFieldsAtStart > 0 ||
      block.read.atoms.some(({ fields }) =>
        fields.some(
          (index) =>
            out.story.state.fields[index]?.paragraphOrdinal !==
            block.read?.ordinal,
        ),
      )
    ) {
      issues.push({
        code: "field-unbalanced",
        message: `Block ${blockId} sits inside a field of the template.`,
        blockId,
      });
    }
    if (text === "mismatch") {
      issues.push({
        code: "text-mismatch",
        message: `Block ${blockId} does not read back as its source text, links, breaks and objects.`,
        blockId,
      });
    }
    // Signatures are per unit of text; they only compare over the same text.
    if (emphasis === "mismatch" && text === "identical") {
      issues.push({
        code: "emphasis-mismatch",
        message: `Block ${blockId} lost or gained emphasis or highlight.`,
        blockId,
      });
    }
    comparisons.set(paragraph, { text, emphasis, output: block });
  }
  for (const block of source.story.order) {
    if (block.read !== null) {
      tally(
        spans.source,
        piecesContent(
          paragraphPieces({
            read: block.read,
            context: source,
            origin: "document",
            fields: source.story.state.fields,
          }),
        ).signatures,
      );
    }
  }

  // Notes the placed text refers to.
  const noteCounts = (
    kind: "footnoteReference" | "endnoteReference",
    sourceNotes: AnalyzedNotes,
    outputNotes: AnalyzedNotes,
  ): TemplatePlanNoteCounts => {
    const origin: StoryOrigin =
      kind === "footnoteReference" ? "footnotes" : "endnotes";
    const counts: TemplatePlanNoteCounts = {
      source: sourceNotes.notes.length,
      expected: 0,
      output: 0,
      identical: 0,
      mismatch: 0,
    };
    for (const note of sourceNotes.notes) {
      tally(
        spans.source,
        sourceNoteContent({
          note,
          source,
          origin,
          fields: sourceNotes.state.fields,
          resources: sourceResources,
        }).signatures,
      );
    }
    // The output note each source note was written as.
    const outputIdOf = new Map<string, string | undefined>();
    const seenOutput = new Set<string>();
    for (const paragraph of resolvedParagraphs) {
      const blockId = paragraph.parts.at(0)?.block.block.id ?? "";
      const expectedIds = noteReferenceIds(
        segmentPieces(paragraph.segments).map(({ kind: pieceKind, atom }) => ({
          kind: pieceKind,
          element: atom?.element ?? null,
        })),
        kind,
      );
      const outputIds = noteReferenceIds(
        outputOf.get(paragraph)?.read?.atoms ?? [],
        kind,
      );
      for (const [index, sourceId] of expectedIds.entries()) {
        const outputId = outputIds[index];
        if (outputIdOf.has(sourceId)) {
          // A second reference to one note must lead to the same note.
          if (outputIdOf.get(sourceId) !== outputId) {
            issues.push({
              code: "note-mismatch",
              message: `Two references to note ${sourceId} of the source lead to different notes.`,
              blockId,
            });
          }
          continue;
        }
        outputIdOf.set(sourceId, outputId);
        counts.expected += 1;
        const sourceNote = sourceNotes.byId.get(sourceId);
        const outputNote =
          outputId === undefined ? undefined : outputNotes.byId.get(outputId);
        const reused = outputId !== undefined && seenOutput.has(outputId);
        if (outputId !== undefined && !reused) {
          seenOutput.add(outputId);
          counts.output += 1;
        }
        const expected =
          sourceNote === undefined
            ? null
            : sourceNoteContent({
                note: sourceNote,
                source,
                origin,
                fields: sourceNotes.state.fields,
                resources: sourceResources,
              });
        const actual =
          outputNote === undefined
            ? null
            : outputNoteContent({
                note: outputNote,
                origin,
                resources: outputResources,
                styles: characterStyles,
              });
        if (expected !== null) {
          tally(spans.expected, expected.signatures);
        }
        if (actual !== null) {
          tally(spans.output, actual.signatures);
        }
        if (
          !reused &&
          expected !== null &&
          actual !== null &&
          outputNote?.hasOtherBlocks === false &&
          sameReading(actual, expected) &&
          sameList(actual.signatures, expected.signatures)
        ) {
          counts.identical += 1;
          continue;
        }
        counts.mismatch += 1;
        issues.push({
          code: "note-mismatch",
          message: `Note ${sourceId} of the source does not read back identical.`,
          blockId,
        });
      }
    }
    return counts;
  };
  const footnotes = noteCounts(
    "footnoteReference",
    source.footnotes,
    out.footnotes,
  );
  const endnotes = noteCounts(
    "endnoteReference",
    source.endnotes,
    out.endnotes,
  );

  // References: the field, its bookmark on the target, and its cached text.
  const outputBookmarks = new Map<string, AnalyzedBlock>();
  for (const block of out.story.order) {
    for (const { name } of block.read?.bookmarks ?? []) {
      outputBookmarks.set(name, block);
    }
  }
  const referenceResults: TemplatePlanReferenceResult[] = [];
  const referencesOf = new Map<ResolvedParagraph, ResolvedPlan["references"]>();
  for (const entry of resolved.references) {
    referencesOf.set(entry.paragraph, [
      ...(referencesOf.get(entry.paragraph) ?? []),
      entry,
    ]);
  }
  for (const [paragraph, entries] of referencesOf) {
    const block = outputOf.get(paragraph);
    const fields = (block?.read?.fieldIndices ?? [])
      .map((index) => out.story.state.fields[index])
      .filter((field) => field?.kind === "REF");
    // Fields come in the order the segments were written, each over the
    // text its segment stands at.
    let offset = 0;
    const ordered = paragraph.segments.flatMap((segment) => {
      const at = offset;
      offset += segmentLength(segment);
      const entry =
        segment.kind === "reference"
          ? entries.find(({ reference }) => reference === segment.reference)
          : undefined;
      return entry === undefined ? [] : [{ entry, at }];
    });
    for (const [index, { entry, at }] of ordered.entries()) {
      const field = fields[index];
      const bookmark = resolved.bookmarks.get(entry.target)?.name ?? "";
      const targetBlock = outputOf.get(entry.target);
      const hasNumberSwitch = NUMBER_SWITCH.test(entry.reference.instruction);
      const usesNumber = /\\n\b/u.test(entry.reference.instruction);
      let predictedText: string | null = null;
      if (targetBlock !== undefined) {
        if (!hasNumberSwitch) {
          predictedText = targetBlock.block.text;
        } else if (usesNumber) {
          predictedText = /\\t\b/u.test(entry.reference.instruction)
            ? (targetBlock.label?.numberOnly ?? null)
            : (targetBlock.label?.paragraphNumber ?? null);
        }
      }
      let status: TemplatePlanReferenceResult["status"] = "ok";
      if (
        field === undefined ||
        fieldBookmarkOf(field.instruction) !== bookmark ||
        field.instruction.trim() !== entry.instruction.trim() ||
        // A field over other characters would replace those when it updates.
        // Offsets only compare over the same text; other text is an issue
        // of its own.
        (comparisons.get(paragraph)?.text === "identical" &&
          (field.paragraphOrdinal !== block?.read?.ordinal ||
            field.resultStart !== at))
      ) {
        status = "missing-field";
      } else if (
        targetBlock === undefined ||
        outputBookmarks.get(bookmark) !== targetBlock ||
        // The bookmark is the target's whole content, no more and no less.
        !bookmarkEncloses(targetBlock.read, bookmark)
      ) {
        status = "missing-bookmark";
      } else if (field.result !== entry.reference.expectedText) {
        status = "cache-mismatch";
      }
      if (status !== "ok") {
        issues.push({
          code: "reference-broken",
          message: `A reference in block ${entry.reference.blockId} to ${entry.reference.targetBlockId} is ${status}.`,
          blockId: entry.reference.blockId,
        });
      }
      referenceResults.push({
        blockId: entry.reference.blockId,
        targetBlockId: entry.reference.targetBlockId,
        bookmark,
        instruction: entry.reference.instruction,
        expectedText: entry.reference.expectedText,
        cachedText: field?.result ?? null,
        replacedText: entry.replacedText,
        predictedText,
        status,
      });
    }
  }

  // The lists the plan makes and the levels it patches, as the output's
  // numbering part states them.
  const outputNumbering =
    out.numberingPath === null || out.numberingXml === null
      ? null
      : readRawPart(out.numberingPath, out.numberingXml);
  const numberingMismatch = (message: string): void => {
    issues.push({
      code: "numbering-mismatch",
      message,
      ...(out.numberingPath === null ? {} : { part: out.numberingPath }),
    });
  };
  for (const mint of resolved.mints) {
    const instance = wChildren(outputNumbering?.root, "num").find(
      (element) => integerOrNull(wAttr(element, "numId")) === mint.numId,
    );
    if (
      instance === undefined ||
      integerOrNull(wChildVal(instance, "abstractNumId")) !== mint.abstractNumId
    ) {
      numberingMismatch(
        `List ${mint.numId} is not an instance of list definition ${mint.abstractNumId}.`,
      );
      continue;
    }
    for (const [level, start] of mint.starts) {
      const override = wChildren(instance, "lvlOverride").find(
        (element) => integerOrNull(wAttr(element, "ilvl")) === level,
      );
      if (integerOrNull(wChildVal(override, "startOverride")) !== start) {
        numberingMismatch(
          `Level ${level} of list ${mint.numId} does not start at ${start}.`,
        );
      }
    }
  }
  for (const patch of plan.numberingPatches ?? []) {
    const definition = wChildren(outputNumbering?.root, "abstractNum").find(
      (element) =>
        integerOrNull(wAttr(element, "abstractNumId")) === patch.abstractNumId,
    );
    const level = wChildren(definition, "lvl").find(
      (element) => integerOrNull(wAttr(element, "ilvl")) === patch.ilvl,
    );
    const planned = {
      start: patch.start === undefined ? undefined : String(patch.start),
      numFmt: patch.numFmt,
      suff: patch.suff,
      lvlText: patch.lvlText,
    };
    for (const [name, value] of Object.entries(planned)) {
      if (value !== undefined && wChildVal(level, name) !== value) {
        numberingMismatch(
          `Level ${patch.ilvl} of list definition ${patch.abstractNumId} does not carry the planned ${name}.`,
        );
      }
    }
    const indent = wChild(wChild(level, "pPr"), "ind");
    for (const [name, value] of Object.entries(patch.indent ?? {})) {
      if (twipsOrNull(wAttr(indent, name)) !== value) {
        numberingMismatch(
          `Level ${patch.ilvl} of list definition ${patch.abstractNumId} does not carry the planned ${name} indent.`,
        );
      }
    }
    // A hanging and a first-line indent exclude each other: the one the
    // plan states must not stand beside the other.
    const excluded = [
      ["hanging", "firstLine"],
      ["firstLine", "hanging"],
    ] as const;
    for (const [stated, other] of excluded) {
      if (
        patch.indent?.[stated] !== undefined &&
        wAttr(indent, other) !== null
      ) {
        numberingMismatch(
          `Level ${patch.ilvl} of list definition ${patch.abstractNumId} keeps a ${other} indent beside the planned ${stated} one.`,
        );
      }
    }
  }

  // Integrity of every id the output's stories use: the main document and
  // the note parts the writer adds to.
  const bookmarkNames = new Map<string, number>();
  const bump = (map: Map<string, number>, key: string): void => {
    map.set(key, (map.get(key) ?? 0) + 1);
  };
  const checkedParts: { origin: StoryOrigin; part: RawPart }[] = [
    { origin: "document", part: out.document },
  ];
  for (const [origin, notes] of [
    ["footnotes", out.footnotes],
    ["endnotes", out.endnotes],
  ] as const) {
    if (notes.part !== null) {
      checkedParts.push({ origin, part: notes.part });
    }
  }
  for (const { origin, part } of checkedParts) {
    const bookmarkStarts = new Map<string, number>();
    const bookmarkEnds = new Map<string, number>();
    const endsBeforeStart = new Set<string>();
    const dangling = {
      relationships: new Set<string>(),
      styles: new Set<string>(),
      numbering: new Set<string>(),
    };
    walkElements(part.root, (element) => {
      const name = wName(element);
      if (name === "bookmarkStart") {
        bump(bookmarkStarts, wAttr(element, "id") ?? "");
        bump(bookmarkNames, wAttr(element, "name") ?? "");
      } else if (name === "bookmarkEnd") {
        const id = wAttr(element, "id") ?? "";
        if (!bookmarkStarts.has(id)) {
          endsBeforeStart.add(id);
        }
        bump(bookmarkEnds, id);
      } else if (
        name === "pStyle" ||
        name === "rStyle" ||
        name === "tblStyle"
      ) {
        const styleId = wAttr(element, "val");
        if (styleId !== null && !out.styles.has(styleId)) {
          dangling.styles.add(styleId);
        }
      } else if (name === "numId") {
        const numId = integerOrNull(wAttr(element, "val")) ?? undefined;
        if (
          isNumberingReference(numId) &&
          out.numbering?.hasNumbering(numId) !== true
        ) {
          dangling.numbering.add(String(numId));
        }
      }
    });
    for (const id of relationshipIdsOf(part.root)) {
      const relationship = outputResources.relationships[origin].get(id);
      if (
        relationship === undefined ||
        (relationship.targetMode !== "External" &&
          !out.pkg.paths.has(
            outputResources.resolveTarget(origin, relationship.target),
          ))
      ) {
        dangling.relationships.add(id);
      }
    }
    for (const [id, count] of bookmarkStarts) {
      if (count !== 1 || bookmarkEnds.get(id) !== 1) {
        issues.push({
          code: "bookmark-dangling",
          message: `Bookmark id ${id} is not one start and one end.`,
          part: part.path,
        });
      }
    }
    for (const id of endsBeforeStart) {
      issues.push({
        code: "bookmark-dangling",
        message: `Bookmark ${id} ends before it starts.`,
        part: part.path,
      });
    }
    for (const id of dangling.relationships) {
      issues.push({
        code: "relationship-dangling",
        message: `Relationship ${id} of ${part.path} leads nowhere.`,
        part: part.path,
      });
    }
    for (const styleId of dangling.styles) {
      issues.push({
        code: "style-dangling",
        message: `Style ${styleId} is used and not defined.`,
        part: part.path,
      });
    }
    for (const numId of dangling.numbering) {
      issues.push({
        code: "numbering-dangling",
        message: `List ${numId} is used and not defined.`,
        part: part.path,
      });
    }
  }
  for (const [name, count] of bookmarkNames) {
    if (count > 1) {
      issues.push({
        code: "bookmark-dangling",
        message: `Bookmark name ${name} is used ${count} times.`,
        part: out.pkg.documentPath,
      });
    }
  }
  for (const block of out.story.order) {
    const atoms = block.read?.atoms ?? [];
    for (const [kind, notes] of [
      ["footnoteReference", out.footnotes],
      ["endnoteReference", out.endnotes],
    ] as const) {
      for (const id of noteReferenceIds(atoms, kind)) {
        // A reference leads to a note with content, not to a separator.
        if (!notes.byId.has(id)) {
          issues.push({
            code: "note-dangling",
            message: `Block ${block.block.id} refers to the missing note ${id}.`,
            blockId: block.block.id,
          });
        }
      }
    }
  }
  for (const { path, state } of [
    { path: out.pkg.documentPath, state: out.story.state },
    { path: out.footnotes.path, state: out.footnotes.state },
    { path: out.endnotes.path, state: out.endnotes.state },
  ]) {
    if (state.open.length > 0 || state.strayFieldEnds > 0) {
      issues.push({
        code: "field-unbalanced",
        message:
          "Complex field marks do not pair up: a kept template field was split.",
        ...(path === null ? {} : { part: path }),
      });
    }
  }

  // Fields the output still carries, and whether the planned freeze took.
  const outputParts: { part: string; fields: StoryField[]; text: string }[] =
    [];
  const outputHeadersAndFooters = await readHeaderFooterParts(out);
  if (Result.isError(outputHeadersAndFooters)) {
    return outputHeadersAndFooters;
  }
  for (const part of outputHeadersAndFooters.value) {
    outputParts.push({
      part: part.path,
      fields: partFields(part),
      text: partText(part),
    });
  }
  for (const notes of [out.footnotes, out.endnotes]) {
    if (notes.part !== null) {
      outputParts.push({
        part: notes.part.path,
        fields: notes.state.fields,
        text: partText(notes.part),
      });
    }
  }
  // Fields of kept template blocks; those of a filled slot are the plan's.
  const writtenOrdinals = new Set(
    [...outputOf.values()].map((block) => block.read?.ordinal ?? -1),
  );
  const templateFields = out.story.state.fields.filter(
    (field) =>
      !writtenOrdinals.has(field.paragraphOrdinal) &&
      templateOutputIds.has(
        out.story.paragraphs[field.paragraphOrdinal]?.topLevelId ?? "",
      ),
  );
  const leftoverFields = [
    ...countByKind(out.pkg.documentPath, templateFields),
    ...outputParts.flatMap(({ part, fields }) => countByKind(part, fields)),
  ];
  const settingsXml = await readPartText(out.pkg, out.pkg.partPath("settings"));
  if (Result.isError(settingsXml)) {
    return settingsXml;
  }
  const updatesOnOpen = parseSettings(settingsXml.value).updateFields === true;
  const allFields = [
    ...out.story.state.fields,
    ...outputParts.flatMap(({ fields }) => fields),
  ];
  let nestedUnfrozen = 0;
  for (const field of allFields) {
    if (updatesOnOpen && PROMPTING_FIELD_KINDS.has(field.kind)) {
      issues.push({
        code: "prompting-field-left",
        message: `A ${field.kind} field would prompt when fields update on open.`,
      });
    } else if (resolved.freezeKinds.has(field.kind) && field.nestedInCode) {
      // Part of another field's code: freezing it would rewrite that code.
      nestedUnfrozen += 1;
    } else if (resolved.freezeKinds.has(field.kind)) {
      issues.push({
        code: "field-not-frozen",
        message: `A ${field.kind} field the plan freezes is still live.`,
      });
    }
  }
  if (nestedUnfrozen > 0) {
    warnings.push(
      `${nestedUnfrozen} field(s) of a frozen kind sit inside another field's code and stay live.`,
    );
  }
  const updateOnOpen = plan.fields?.updateOnOpen;
  if (updateOnOpen !== undefined && updateOnOpen !== updatesOnOpen) {
    issues.push({
      code: "update-fields-mismatch",
      message: "The update-on-open setting is not what the plan asks for.",
    });
  }
  for (const field of templateFields) {
    const bookmark = BOOKMARK_FIELD_KINDS.has(field.kind)
      ? fieldBookmarkOf(field.instruction)
      : null;
    if (bookmark !== null && !bookmarkNames.has(bookmark)) {
      warnings.push(
        `A kept ${field.kind} field points at the bookmark ${bookmark}, which the output does not hold; its cached text is stale until fields update.`,
      );
    }
  }
  const keptToc = templateFields.filter(({ kind }) => kind === "TOC").length;
  if (keptToc > 0) {
    warnings.push(
      `${keptToc} table of contents field(s) keep the template's cached entries until fields update.`,
    );
  }

  // Frozen fields, counted against the template the plan started from.
  const keptTemplateIds = new Set(
    resolved.keptTemplateBlocks.map(({ block }) => block.id),
  );
  const frozenFields: TemplatePlanFieldLeftover[] = [];
  if (resolved.freezeKinds.size > 0) {
    const slotOrdinals = new Set(
      resolved.nestedSlots.map(
        ({ templateBlock }) => templateBlock.read?.ordinal ?? -1,
      ),
    );
    const frozenIn = (fields: readonly StoryField[]): StoryField[] =>
      fields.filter(
        ({ kind, nestedInCode }) =>
          resolved.freezeKinds.has(kind) && !nestedInCode,
      );
    frozenFields.push(
      ...countByKind(
        template.pkg.documentPath,
        frozenIn(template.story.state.fields).filter(
          (field) =>
            !slotOrdinals.has(field.paragraphOrdinal) &&
            keptTemplateIds.has(
              template.story.paragraphs[field.paragraphOrdinal]?.topLevelId ??
                "",
            ),
        ),
      ),
    );
    for (const notes of [template.footnotes, template.endnotes]) {
      if (notes.path !== null) {
        frozenFields.push(
          ...countByKind(notes.path, frozenIn(notes.state.fields)),
        );
      }
    }
    const templateHeadersAndFooters = await readHeaderFooterParts(template);
    if (Result.isError(templateHeadersAndFooters)) {
      return templateHeadersAndFooters;
    }
    for (const part of templateHeadersAndFooters.value) {
      frozenFields.push(...countByKind(part.path, frozenIn(partFields(part))));
    }
  }

  const filler: TemplatePlanReport["leftovers"]["filler"] = [];
  const documentText = out.story.order
    .flatMap(({ read }) => (read === null ? [] : [read.text]))
    .join("\n");
  for (const text of options?.fillerStrings ?? []) {
    for (const { part, text: haystack } of [
      { part: out.pkg.documentPath, text: documentText },
      ...outputParts,
    ]) {
      const count = occurrences(haystack, text);
      if (count > 0) {
        filler.push({ text, part, count });
      }
    }
  }
  if (filler.length > 0) {
    warnings.push(
      `${filler.length} filler string(s) of the template are still in the output.`,
    );
  }

  // One result per source block.
  const blocks: TemplatePlanBlockResult[] = [];
  const counts: TemplatePlanReport["counts"] = {
    sourceBlocks: source.story.order.length,
    placed: 0,
    merged: 0,
    omitted: 0,
    templateBlocksKept: resolved.keptTemplateBlocks.length,
    textIdentical: 0,
    textMismatch: 0,
    emphasisMismatch: 0,
    labelChanged: 0,
    labelLost: 0,
  };
  const tables = new Map<string, TemplatePlanBlockResult>();
  for (const block of source.story.order) {
    const { id } = block.block;
    const omitReason =
      resolved.omitted.get(block.topLevelId) ?? resolved.omitted.get(id);
    if (omitReason !== undefined) {
      counts.omitted += 1;
      blocks.push({
        blockId: id,
        outputBlockId: null,
        disposition: "omitted",
        text: "omitted",
        emphasis: "omitted",
        sourceLabel: block.block.label,
        outputLabel: null,
        omitReason,
      });
      continue;
    }
    if (block.read === null) {
      const result: TemplatePlanBlockResult = {
        blockId: id,
        outputBlockId: tableOutputs.get(id) ?? null,
        disposition: block.parentId === null ? "placed" : "nested",
        text: "identical",
        emphasis: "identical",
        sourceLabel: null,
        outputLabel: null,
      };
      counts.placed += 1;
      tables.set(id, result);
      blocks.push(result);
      continue;
    }
    const paragraph = resolved.byBlockId.get(id);
    const comparison =
      paragraph === undefined ? undefined : comparisons.get(paragraph);
    const partIndex =
      paragraph?.parts.findIndex((part) => part.block === block) ?? -1;
    const part = paragraph?.parts[partIndex];
    const leads = partIndex === 0;
    const strippedLabel = part?.strippedPrefix.trim() ?? "";
    const sourceLabel =
      block.block.label ?? (strippedLabel === "" ? null : strippedLabel);
    const outputLabel = leads
      ? (comparison?.output?.block.label ?? null)
      : null;
    let disposition: TemplatePlanBlockResult["disposition"] = "placed";
    if (!leads) {
      disposition = "merged";
    } else if (paragraph?.placement === "slot") {
      disposition = "slot";
    } else if (paragraph?.placement === "nested") {
      disposition = "nested";
    }
    const text = comparison?.text ?? "mismatch";
    const emphasis = comparison?.emphasis ?? "mismatch";
    if (disposition === "merged") {
      counts.merged += 1;
    } else {
      counts.placed += 1;
    }
    if (text === "identical") {
      counts.textIdentical += 1;
    } else {
      counts.textMismatch += 1;
    }
    if (emphasis === "mismatch") {
      counts.emphasisMismatch += 1;
    }
    if (sourceLabel !== null && outputLabel === null) {
      counts.labelLost += 1;
    } else if (
      sourceLabel !== null &&
      outputLabel !== null &&
      sourceLabel.trim() !== outputLabel.trim()
    ) {
      counts.labelChanged += 1;
    }
    // A mismatch shows on every table around the paragraph.
    let tableId = block.parentId;
    while (tableId !== null) {
      const table = tables.get(tableId);
      if (table !== undefined && text === "mismatch") {
        table.text = "mismatch";
      }
      if (table !== undefined && emphasis === "mismatch") {
        table.emphasis = "mismatch";
      }
      tableId = source.story.byId.get(tableId)?.parentId ?? null;
    }
    blocks.push({
      blockId: id,
      outputBlockId: comparison?.output?.block.id ?? null,
      disposition,
      text,
      emphasis,
      ...(part === undefined || part.strippedPrefix === ""
        ? {}
        : { strippedPrefix: part.strippedPrefix }),
      sourceLabel,
      outputLabel,
    });
  }
  if (counts.labelLost > 0) {
    warnings.push(
      `${counts.labelLost} block(s) showed a number or label in the source and render none.`,
    );
  }
  if (counts.labelChanged > 0) {
    warnings.push(
      `${counts.labelChanged} block(s) render a label that differs from the source's.`,
    );
  }

  // What the transfer leaves behind by design.
  const placedReads = resolvedParagraphs.flatMap(({ parts }) =>
    parts.flatMap(({ block }) => (block.read === null ? [] : [block.read])),
  );
  const droppedFields = new Map<string, number>();
  let internalHyperlinks = 0;
  for (const read of placedReads) {
    for (const index of read.fieldIndices) {
      const field = source.story.state.fields[index];
      // A HYPERLINK field with an address is carried over as a link.
      if (field !== undefined && hyperlinkFieldLink(field) === null) {
        droppedFields.set(field.kind, (droppedFields.get(field.kind) ?? 0) + 1);
      }
    }
    // Links the output cannot carry: to a bookmark or a part of the source.
    let previous: AtomLink | null = null;
    for (const { link } of read.atoms) {
      const lost =
        link !== null &&
        linkAddress(link, sourceResources.relationships.document) === null;
      if (lost && link !== previous) {
        internalHyperlinks += 1;
      }
      previous = lost ? link : null;
    }
  }
  // Direct paragraph formatting stays behind, and section ends with it.
  const sectionBreaks = source.story.paragraphs.filter(
    (block) => block.read?.hasSectionBreak === true,
  ).length;
  const omittedBlocks = source.story.order.filter(
    (block) =>
      resolved.omitted.has(block.topLevelId) ||
      resolved.omitted.has(block.block.id),
  );
  let omittedBreaks = 0;
  for (const block of omittedBlocks) {
    for (const atom of block.read?.atoms ?? []) {
      if (atom.kind === "pageBreak") {
        omittedBreaks += 1;
      }
    }
  }
  const omittedLabels = omittedBlocks.filter(
    ({ block }) => block.label !== null,
  ).length;
  if (omittedBreaks > 0) {
    warnings.push(
      `${omittedBreaks} page or column break(s) of the source sit in omitted blocks.`,
    );
  }
  if (omittedLabels > 0) {
    warnings.push(
      `${omittedLabels} omitted block(s) showed a number or label in the source.`,
    );
  }
  if (source.story.state.comments > 0) {
    warnings.push(
      `${source.story.state.comments} comment(s) of the source are not carried over.`,
    );
  }
  if (droppedFields.size > 0) {
    warnings.push(
      "Fields of the source are carried over as their cached text.",
    );
  }
  if (internalHyperlinks > 0) {
    warnings.push(
      `${internalHyperlinks} link(s) to bookmarks of the source are carried over as text.`,
    );
  }
  if (sectionBreaks > 0) {
    warnings.push(
      `${sectionBreaks} section break(s) of the source are not carried over.`,
    );
  }

  const [conformance, templateConformance] = await Promise.all([
    validateDocxConformance(output),
    validateDocxConformance(template.pkg.bytes),
  ]);
  if (
    CONFORMANCE_RANK[conformance.status] <
    CONFORMANCE_RANK[templateConformance.status]
  ) {
    issues.push({
      code: "conformance",
      message: `The output package is ${conformance.status}; the template is ${templateConformance.status}.`,
    });
  }

  const predicted = referenceResults.filter(
    ({ predictedText }) => predictedText !== null,
  );
  const predictionMismatch = predicted.filter(
    ({ predictedText, expectedText }) => predictedText !== expectedText,
  ).length;
  if (predictionMismatch > 0) {
    warnings.push(
      `${predictionMismatch} reference(s) cache a text that differs from the target's rendered label.`,
    );
  }
  const textChanging = referenceResults.filter(
    ({ expectedText, replacedText }) => expectedText !== replacedText,
  ).length;
  if (textChanging > 0) {
    warnings.push(
      `${textChanging} reference(s) show a text that differs from the source's wording.`,
    );
  }

  let updateFields: TemplatePlanReport["updateFields"] = "unchanged";
  if (updateOnOpen === true) {
    updateFields = "set";
  } else if (updateOnOpen === false) {
    updateFields = "cleared";
  }
  return Result.ok({
    version: 1,
    ok: issues.length === 0,
    sourceSha256: source.manifest.sha256,
    templateSha256: template.manifest.sha256,
    outputSha256: out.manifest.sha256,
    blocks,
    counts,
    footnotes,
    endnotes,
    emphasisSpans: {
      source: spans.source.emphasis,
      expected: spans.expected.emphasis,
      output: spans.output.emphasis,
    },
    highlightSpans: {
      source: spans.source.highlight,
      expected: spans.expected.highlight,
      output: spans.output.highlight,
    },
    references: {
      planned: resolved.references.length,
      ok: referenceResults.filter(({ status }) => status === "ok").length,
      textChanging,
      predictionMismatch,
      results: referenceResults,
    },
    leftovers: {
      fields: leftoverFields,
      frozenFields,
      filler,
      templateBlockIds: resolved.keptTemplateBlocks.map(
        ({ block }) => block.id,
      ),
    },
    dropped: {
      comments: source.story.state.comments,
      sourceFields: [...droppedFields].map(([kind, count]) => ({
        kind,
        count,
      })),
      internalHyperlinks,
      sectionBreaks,
    },
    updateFields,
    conformance: {
      status: conformance.status,
      templateStatus: templateConformance.status,
      issues: [...conformance.issues],
    },
    issues,
    warnings,
  });
};

const verifyUnguarded = async ({
  source,
  template,
  output,
  plan,
  options,
}: VerifyTemplatePlanOutputArgs): Promise<
  Result<TemplatePlanReport, TemplatePlanError>
> => {
  const analyzedSource = await analyzeDocument(source);
  if (Result.isError(analyzedSource)) {
    return analyzedSource;
  }
  const analyzedTemplate = await analyzeDocument(template);
  if (Result.isError(analyzedTemplate)) {
    return analyzedTemplate;
  }
  const resolved = resolvePlan(
    analyzedSource.value,
    analyzedTemplate.value,
    plan,
  );
  if (Result.isError(resolved)) {
    return resolved;
  }
  return await reportTemplatePlanOutput({
    source: analyzedSource.value,
    template: analyzedTemplate.value,
    output,
    plan,
    resolved: resolved.value,
    options,
  });
};

/**
 * Check `output` against `source`, `template` and `plan` without trusting
 * whoever wrote it: every expectation is derived again from the inputs.
 * A refused plan is an error; a failed check is a report with `ok: false`.
 */
export const verifyTemplatePlanOutput = async (
  input: VerifyTemplatePlanOutputArgs,
): Promise<Result<TemplatePlanReport, TemplatePlanError>> =>
  await guarded(async () => await verifyUnguarded(input));
