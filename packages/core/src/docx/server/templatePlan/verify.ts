/**
 * Verification of an applied template plan. The output package is read back
 * with the same reader that analysed the source and compared, block by block,
 * against what the plan resolves to from the source alone: text, stated
 * emphasis, notes, references and the integrity of every id the body uses.
 */
import { Result } from "better-result";

import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";

import { parseSettings } from "../../settingsParser";
import {
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  type XmlElement,
} from "../../xmlParser";
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
  emphasisSignature,
  readDirectEmphasis,
  signatureHighlight,
  signatureWithoutHighlight,
} from "./emphasis";
import { headerFooterPaths, partFields } from "./inspect";
import { readPartText } from "./package";
import {
  paragraphPieces,
  type PieceOrigin,
  piecesContent,
  type ResolvedParagraph,
  type ResolvedPlan,
  segmentPieces,
} from "./plan";
import {
  createStoryState,
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
} from "./types";
import {
  integerOrNull,
  isW,
  type RawPart,
  readRawPart,
  walkElements,
  wAttr,
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

/** What a paragraph of the output states, unit by unit, read off its runs. */
const outputSignatures = (read: ParagraphRead): string[] => {
  const signatures: string[] = [];
  for (const atom of read.atoms) {
    const signature = NOTE_REFERENCE_KINDS.has(atom.kind)
      ? ""
      : emphasisSignature(readDirectEmphasis(atom.runProperties));
    for (let unit = 0; unit < atom.text.length; unit += 1) {
      signatures.push(signature);
    }
  }
  return signatures;
};

const sameSignatures = (
  left: readonly string[],
  right: readonly string[],
): boolean =>
  left.length === right.length &&
  left.every((signature, index) => signature === right[index]);

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

type NoteContent = { text: string; signatures: string[] };

const sourceNoteContent = (
  note: AnalyzedNote,
  source: AnalyzedDocument,
  origin: PieceOrigin,
  fields: readonly StoryField[],
): NoteContent => {
  const content: NoteContent = { text: "", signatures: [] };
  for (const [index, read] of note.paragraphs.entries()) {
    const paragraph = piecesContent(
      paragraphPieces({ read, context: source, origin, fields }),
    );
    if (index > 0) {
      content.text += "\n";
      content.signatures.push("");
    }
    content.text += paragraph.text;
    content.signatures.push(...paragraph.signatures);
  }
  return content;
};

const outputNoteContent = (note: AnalyzedNote): NoteContent => {
  const content: NoteContent = { text: "", signatures: [] };
  for (const [index, read] of note.paragraphs.entries()) {
    if (index > 0) {
      content.text += "\n";
      content.signatures.push("");
    }
    content.text += read.text;
    content.signatures.push(...outputSignatures(read));
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
    });
  };
  if (isW(part.root, "footnotes") || isW(part.root, "endnotes")) {
    for (const note of part.root.elements ?? []) {
      if (note.type === "element") {
        readUnder(note);
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
  const warnings: string[] = [...out.story.warnings];

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
        const sourceNested = nestedParagraphs(
          source.story,
          item.block.block.id,
        );
        const outputNested = nestedParagraphs(out.story, block.block.id);
        if (sourceNested.length !== outputNested.length) {
          missing(
            `Source table ${item.block.block.id} has ${sourceNested.length} paragraph(s), the output ${outputNested.length}.`,
          );
        }
        for (const [index, nested] of sourceNested.entries()) {
          assign(item.cells.get(nested.element), outputNested[index]);
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
    const signatures = outputSignatures(block.read);
    tally(spans.output, signatures);
    const text =
      block.read.text === paragraph.text &&
      block.read.inlineObjects === paragraph.inlineObjects
        ? "identical"
        : "mismatch";
    const emphasis = sameSignatures(signatures, paragraph.signatures)
      ? "identical"
      : "mismatch";
    if (text === "mismatch") {
      issues.push({
        code: "text-mismatch",
        message: `Block ${blockId} does not read back as its source text.`,
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
    const origin: PieceOrigin =
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
        sourceNoteContent(note, source, origin, sourceNotes.state.fields)
          .signatures,
      );
    }
    const seenSource = new Set<string>();
    const seenOutput = new Set<string>();
    for (const paragraph of resolvedParagraphs) {
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
        if (seenSource.has(sourceId)) {
          continue;
        }
        seenSource.add(sourceId);
        counts.expected += 1;
        const sourceNote = sourceNotes.byId.get(sourceId);
        const outputId = outputIds[index];
        const outputNote =
          outputId === undefined ? undefined : outputNotes.byId.get(outputId);
        if (outputId !== undefined && !seenOutput.has(outputId)) {
          seenOutput.add(outputId);
          counts.output += 1;
        }
        const expected =
          sourceNote === undefined
            ? { text: "", signatures: [] }
            : sourceNoteContent(
                sourceNote,
                source,
                origin,
                sourceNotes.state.fields,
              );
        tally(spans.expected, expected.signatures);
        const actual =
          outputNote === undefined ? null : outputNoteContent(outputNote);
        if (actual !== null) {
          tally(spans.output, actual.signatures);
        }
        if (
          actual !== null &&
          actual.text === expected.text &&
          sameSignatures(actual.signatures, expected.signatures)
        ) {
          counts.identical += 1;
          continue;
        }
        counts.mismatch += 1;
        issues.push({
          code: "note-mismatch",
          message: `Note ${sourceId} of the source does not read back identical.`,
          blockId: paragraph.parts.at(0)?.block.block.id ?? "",
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
    // Fields come in the order the segments were written.
    const ordered = paragraph.segments.flatMap((segment) => {
      const entry =
        segment.kind === "reference"
          ? entries.find(({ reference }) => reference === segment.reference)
          : undefined;
      return entry === undefined ? [] : [entry];
    });
    for (const [index, entry] of ordered.entries()) {
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
        field.instruction.trim() !== entry.instruction.trim()
      ) {
        status = "missing-field";
      } else if (
        targetBlock === undefined ||
        outputBookmarks.get(bookmark) !== targetBlock
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

  // Integrity of every id the output's main document uses.
  const bookmarkStarts = new Map<string, number>();
  const bookmarkEnds = new Map<string, number>();
  const bookmarkNames = new Map<string, number>();
  const bump = (map: Map<string, number>, key: string): void => {
    map.set(key, (map.get(key) ?? 0) + 1);
  };
  const dangling = {
    relationships: new Set<string>(),
    styles: new Set<string>(),
    numbering: new Set<string>(),
  };
  walkElements(out.document.root, (element) => {
    const name = wName(element);
    if (name === "bookmarkStart") {
      bump(bookmarkStarts, wAttr(element, "id") ?? "");
      bump(bookmarkNames, wAttr(element, "name") ?? "");
    } else if (name === "bookmarkEnd") {
      bump(bookmarkEnds, wAttr(element, "id") ?? "");
    } else if (name === "pStyle" || name === "rStyle" || name === "tblStyle") {
      const styleId = wAttr(element, "val");
      if (styleId !== null && !out.styles.has(styleId)) {
        dangling.styles.add(styleId);
      }
    } else if (name === "numId") {
      const numId = integerOrNull(wAttr(element, "val"));
      if (
        numId !== null &&
        numId !== 0 &&
        out.numbering?.hasNumbering(numId) !== true
      ) {
        dangling.numbering.add(String(numId));
      }
    }
    for (const [attribute, value] of Object.entries(element.attributes ?? {})) {
      const colon = attribute.indexOf(":");
      if (
        value === undefined ||
        colon === -1 ||
        attribute.startsWith("xmlns:") ||
        !OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(
          resolveNamespaceUri(
            element.namespaceScope,
            attribute.slice(0, colon),
          ) ?? "",
        )
      ) {
        continue;
      }
      const relationship = out.pkg.documentRelationships.get(String(value));
      if (
        relationship === undefined ||
        (relationship.targetMode !== "External" &&
          !out.pkg.paths.has(
            out.pkg.resolveDocumentTarget(relationship.target),
          ))
      ) {
        dangling.relationships.add(String(value));
      }
    }
  });
  for (const [id, count] of bookmarkStarts) {
    if (count !== 1 || bookmarkEnds.get(id) !== 1) {
      issues.push({
        code: "bookmark-dangling",
        message: `Bookmark id ${id} is not one start and one end.`,
        part: out.pkg.documentPath,
      });
    }
  }
  for (const id of bookmarkEnds.keys()) {
    if (!bookmarkStarts.has(id)) {
      issues.push({
        code: "bookmark-dangling",
        message: `Bookmark end ${id} has no start.`,
        part: out.pkg.documentPath,
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
  for (const id of dangling.relationships) {
    issues.push({
      code: "relationship-dangling",
      message: `Relationship ${id} of the main document leads nowhere.`,
      part: out.pkg.documentPath,
    });
  }
  for (const styleId of dangling.styles) {
    issues.push({
      code: "style-dangling",
      message: `Style ${styleId} is used and not defined.`,
      part: out.pkg.documentPath,
    });
  }
  for (const numId of dangling.numbering) {
    issues.push({
      code: "numbering-dangling",
      message: `List ${numId} is used and not defined.`,
      part: out.pkg.documentPath,
    });
  }
  for (const block of out.story.order) {
    const atoms = block.read?.atoms ?? [];
    for (const [kind, notes] of [
      ["footnoteReference", out.footnotes],
      ["endnoteReference", out.endnotes],
    ] as const) {
      for (const id of noteReferenceIds(atoms, kind)) {
        if (!notes.usedIds.includes(Number.parseInt(id, 10))) {
          issues.push({
            code: "note-dangling",
            message: `Block ${block.block.id} refers to the missing note ${id}.`,
            blockId: block.block.id,
          });
        }
      }
    }
  }
  if (out.story.state.open.length > 0 || out.story.state.strayFieldEnds > 0) {
    issues.push({
      code: "field-unbalanced",
      message:
        "A complex field of the main story is not closed: a kept template field was split.",
      part: out.pkg.documentPath,
    });
  }

  // Fields the output still carries, and whether the planned freeze took.
  const outputParts: { part: string; fields: StoryField[]; text: string }[] =
    [];
  for (const path of headerFooterPaths(out)) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const xml = await readPartText(out.pkg, path);
    if (Result.isError(xml)) {
      return xml;
    }
    const part = xml.value === null ? null : readRawPart(path, xml.value);
    if (part !== null) {
      outputParts.push({
        part: path,
        fields: partFields(part),
        text: partText(part),
      });
    }
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
  const templateFields = out.story.state.fields.filter((field) =>
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
  for (const field of allFields) {
    if (resolved.freezeKinds.has(field.kind)) {
      issues.push({
        code: "field-not-frozen",
        message: `A ${field.kind} field the plan freezes is still live.`,
      });
    } else if (updatesOnOpen && PROMPTING_FIELD_KINDS.has(field.kind)) {
      issues.push({
        code: "prompting-field-left",
        message: `A ${field.kind} field would prompt when fields update on open.`,
      });
    }
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
    frozenFields.push(
      ...countByKind(
        template.pkg.documentPath,
        template.story.state.fields.filter(
          (field) =>
            resolved.freezeKinds.has(field.kind) &&
            keptTemplateIds.has(
              template.story.paragraphs[field.paragraphOrdinal]?.topLevelId ??
                "",
            ),
        ),
      ),
    );
    for (const path of headerFooterPaths(template)) {
      // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
      const xml = await readPartText(template.pkg, path);
      if (Result.isError(xml)) {
        return xml;
      }
      const part = xml.value === null ? null : readRawPart(path, xml.value);
      if (part !== null) {
        frozenFields.push(
          ...countByKind(
            path,
            partFields(part).filter(({ kind }) =>
              resolved.freezeKinds.has(kind),
            ),
          ),
        );
      }
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
        outputBlockId: null,
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
    for (const tableId of [block.parentId, block.topLevelId]) {
      const table = tableId === null ? undefined : tables.get(tableId);
      if (table === undefined) {
        continue;
      }
      if (text === "mismatch") {
        table.text = "mismatch";
      }
      if (emphasis === "mismatch") {
        table.emphasis = "mismatch";
      }
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
  for (const [position, item] of resolved.items.entries()) {
    const table =
      item.kind === "table" ? tables.get(item.block.block.id) : undefined;
    if (table !== undefined) {
      table.outputBlockId = outputTopLevel[position]?.block.id ?? null;
    }
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
      const kind = source.story.state.fields[index]?.kind;
      if (kind !== undefined) {
        droppedFields.set(kind, (droppedFields.get(kind) ?? 0) + 1);
      }
    }
    let previousAnchor: string | null = null;
    for (const { link } of read.atoms) {
      const anchor = link?.kind === "anchor" ? link.anchor : null;
      if (anchor !== null && anchor !== previousAnchor) {
        internalHyperlinks += 1;
      }
      previousAnchor = anchor;
    }
  }
  const sectionBreaks = placedReads.filter(
    ({ hasSectionBreak }) => hasSectionBreak,
  ).length;
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
