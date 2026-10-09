/** Canonical edits retain authored fields; list rendering follows the current numbering. */
import {
  withBodyContent,
  OP_STORIES,
  DOCUMENT_OP_TYPES,
  EMPTY_PROPERTY_SETS,
  type OpStory,
  type DocumentOp,
} from "@stll/docx-core/ops";
import { panic } from "better-result";
import type { Document, DocumentBody, Paragraph, Footnote, Endnote } from "../types/document";
import { computeListRendering, getCachedNumberingMap } from "../docx/numberingParser";
import {
  mergeParagraphNumbering,
  paragraphNumberingReferenceId,
  sameStatedParagraphNumbering,
} from "../docx/numberingReference";
import { computeListMarker, type ComputeListMarkerOptions } from "../docx/listMarkerComputation";
import {
  copyParagraphPropertySource,
  paragraphFormattingWithAuthoredIndentation,
  assignParagraphIndentationProjection,
  sameParagraphIndentationProjection,
} from "../docx/paragraphPropertySource";
import {
  paragraphIndentationFromFormatting,
  withDirectParagraphIndentation,
} from "../prosemirror/paragraphIndentation";
import { listRenderingDefinitionsMatch } from "../prosemirror/conversion/listRenderingDefinition";
import { canonicalJson } from "../utils/canonicalJson";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import { listIndentationProvenancePatch } from "../prosemirror/styles/resolvedStyleAttrs";
import {
  foldListNumberFields,
  unfoldedListNumberContent,
  planParagraphListNumberFold,
} from "../docx/foldedListNumberFields";

type CanonicalListNormalization = { document: Document; inverse: DocumentOp[] };

export const normalizeCanonicalListRendering = (document: Document): CanonicalListNormalization => {
  const inverse: DocumentOp[] = [];
  const styles = createStyleResolver(document.package.styles);
  const numbering =
    document.package.numbering === undefined
      ? null
      : getCachedNumberingMap(document.package.numbering);
  const normalizeBody = <Body extends DocumentBody>(body: Body, story: OpStory): Body => {
    const counters: ComputeListMarkerOptions = {
      numbering,
      listCounters: new Map(),
      abstractCounters: new Map(),
      restartedNumIds: new Set(),
      previousList: { abstractNumId: null, fromStyle: false, numId: null },
    };
    const content = body.content.map((paragraph) => {
      if (paragraph.type !== "paragraph") return paragraph;
      const formatting = paragraph.formatting;
      const numPr = mergeParagraphNumbering(formatting?.numPrFromStyle, formatting?.numPr);
      if (
        paragraph.listRendering === undefined &&
        formatting?.numPrFromStyle === undefined &&
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
      const next: Paragraph = { ...paragraph };
      if (
        formatting?.numPrFromStyle !== undefined &&
        formatting.numPr !== undefined &&
        paragraphNumberingReferenceId(numPr) !==
          paragraphNumberingReferenceId(formatting.numPrFromStyle)
      ) {
        next.formatting = { ...formatting };
        delete next.formatting.numPrFromStyle;
        // Provenance participates in exact formatting history; rendering is recomputed.
        inverse.push({
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story,
          blockId:
            paragraph.paraId ??
            panic("Canonical numbering normalization requires paragraph identities."),
          patch: { numPrFromStyle: formatting.numPrFromStyle },
        });
      }
      if (
        formatting?.numPrFromStyle !== undefined &&
        numPr?.kind === "reference" &&
        !sameStatedParagraphNumbering(formatting.numPr, numPr)
      ) {
        next.formatting = { ...next.formatting, numPr };
        inverse.push({
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story,
          blockId:
            paragraph.paraId ??
            panic("Canonical numbering normalization requires paragraph identities."),
          patch: { numPr: formatting.numPr ?? null },
        });
      }
      const authored = paragraphIndentationFromFormatting(
        paragraphFormattingWithAuthoredIndentation(paragraph),
      );
      const inheritance = listIndentationProvenancePatch({
        direct: authored,
        styleFormatting: styles.resolveParagraphStyle(next.formatting?.styleId).paragraphFormatting,
        numberingSource: next.formatting?.numPrFromStyle === undefined ? "paragraph" : "style",
        numPr:
          numPr?.kind === "reference" ? { numId: numPr.numId, ilvl: numPr.ilvl ?? 0 } : undefined,
        numbering,
      });
      const authoredFormatting = withDirectParagraphIndentation(next.formatting, authored);
      assignParagraphIndentationProjection({
        paragraph: next,
        authored: authoredFormatting,
        inherited: paragraphIndentationFromFormatting(inheritance._resolvedFormatting) ?? {},
      });
      // Structural inverse preconditions capture the fields before rendering
      // normalization; undo its effective indentation before those inverses.
      if (
        canonicalJson(paragraphIndentationFromFormatting(next.formatting)) !==
        canonicalJson(paragraphIndentationFromFormatting(formatting))
      ) {
        inverse.unshift({
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story,
          blockId:
            paragraph.paraId ??
            panic("Canonical numbering normalization requires paragraph identities."),
          patch: {
            indentLeft: formatting?.indentLeft ?? null,
            indentRight: formatting?.indentRight ?? null,
            indentFirstLine: formatting?.indentFirstLine ?? null,
            hangingIndent: formatting?.hangingIndent ?? null,
          },
          whenEmpty: formatting === undefined ? EMPTY_PROPERTY_SETS.OMIT : EMPTY_PROPERTY_SETS.KEEP,
        });
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
        const child = numbering?.getLevel(rendering.numId, rendering.level + 1)?.pPr;
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
      if (
        sameParagraphIndentationProjection({ left: next, right: paragraph }) &&
        canonicalJson(next.formatting) === canonicalJson(formatting) &&
        canonicalJson(next.listRendering) === canonicalJson(cached)
      )
        return paragraph;
      copyParagraphPropertySource(next, paragraph);
      return next;
    });
    return content.every((paragraph, index) => paragraph === body.content[index])
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
      [...parts].map(([id, part]) => [id, normalizeBody(part, { kind, rId: id })]),
    );
    return [...normalized].every(([id, part]) => part === parts.get(id)) ? parts : normalized;
  };
  const normalizeNotes = <Body extends Footnote | Endnote>(
    notes: Body[] | undefined,
    kind: "footnote" | "endnote",
  ) => {
    if (notes === undefined) return notes;
    const normalized = notes.map((note) => normalizeBody(note, { kind, id: note.id }));
    return normalized.every((note, index) => note === notes[index]) ? notes : normalized;
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
