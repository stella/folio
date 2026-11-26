import type { Node as PMNode } from "prosemirror-model";

import {
  buildCleanBlockText,
  type BuildCleanBlockTextOptions,
} from "../ai-edits/clean-text";
import {
  type NoteReferenceIdentity,
  type NoteReferenceLabels,
} from "../ai-edits/note-references";
import { drawingXmlWithoutIdentity } from "../docx/drawingIdNormalization";
import { runFormattingInlineAtomResultText } from "../prosemirror/runFormattingInlineCarriers";
import { canonicalJson } from "../utils/canonicalJson";
import {
  type AlignableAtom,
  sameSupportedAtoms,
} from "./inline-atom-alignment";
import { prepareTargetInlineAtom } from "./inline-atom-resources";

export { MISSING_NOTE_DETAIL } from "./inline-atom-alignment";

export type TextBlock = { node: PMNode; from: number };

export type InlineAtom = AlignableAtom & {
  node: PMNode;
  from: number;
};

/** A footnote or endnote reference read as its marker, over `length` clean-text characters. */
type NoteReferenceAt = NoteReferenceIdentity & {
  offset: number;
  length: number;
  from: number;
};

export type AtomBlock = TextBlock & {
  blockType: string;
  cleanText: string;
  offsets: readonly number[];
  supported: readonly InlineAtom[];
  unsupportedTopology: readonly { offset: number; key: string }[];
  /** Empty unless the block was read with note reference labels. */
  noteReferences: readonly NoteReferenceAt[];
};

/**
 * How a block's note references read. `"sourceText"` is the editor's own text
 * for a reference (its package id); labels read it as the marker (`[^1]`) the
 * comparison planned in, which is the text its operations write for one.
 */
export type NoteReferenceReading = NoteReferenceLabels | "sourceText";

export const isSupportedAtom = (node: PMNode): boolean =>
  node.isInline &&
  node.isAtom &&
  (node.type.name === "field" ||
    node.type.name === "image" ||
    node.type.name === "pageBreakRun");

export const hasSupportedAtom = (doc: PMNode): boolean => {
  let found = false;
  doc.descendants((node) => {
    if (isSupportedAtom(node)) {
      found = true;
      return false;
    }
    return !found;
  });
  return found;
};

export const textBlocksOf = (doc: PMNode): TextBlock[] => {
  const blocks: TextBlock[] = [];
  doc.descendants((node, from) => {
    if (!node.isTextblock) return true;
    blocks.push({ node, from });
    return false;
  });
  return blocks;
};

/**
 * Clean-text offset nearest a document position.
 *
 * `buildCleanBlockText` returns one offset per character plus one for
 * end-of-block, so there is always an offset to land on. The result was typed
 * `number | null` and never was: the caller's refusal branch could not run,
 * while an empty `offsets` would have produced `-1` rather than the `null` the
 * type promised.
 */
const offsetAt = ({
  offsets,
  position,
}: {
  offsets: readonly number[];
  position: number;
}): number => {
  let low = 0;
  let high = offsets.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    const candidate = offsets[middle];
    if (candidate === undefined || candidate < position) low = middle + 1;
    else high = middle;
  }
  return Math.min(low, offsets.length - 1);
};

/**
 * Attributes a conversion mints for itself rather than reads from the package.
 *
 * `toProseDoc` salts text-box anchor ids with a per-load nonce so a pasted span
 * cannot hijack a real anchor. A comparison reads two independent conversions,
 * so a minted handle differs on every run, and a key holding one makes a
 * document differ from itself. Only document facts belong in a key.
 *
 * The guard against a node type gaining a minted attribute without being added
 * here is not this list but `compareIdentity.property.test.ts`, which compares
 * generated packages carrying each atom with themselves.
 */
const CONVERSION_LOCAL_ATTRS: Readonly<Record<string, readonly string[]>> = {
  textBoxAnchor: ["anchorId"],
};

const documentFactAttrs = (node: PMNode): Record<string, unknown> => {
  // OOXML cannot put w:fldSimple inside a revision. A tracked simple field is
  // written as field-character runs, whose parsed fieldKind is "complex" even
  // though its instruction, cached result, state and formatting are unchanged.
  // A synthesized fallback result is different: it was never authored.
  if (node.type.name === "field" || node.type.name === "structuredField") {
    const equivalentFieldForms =
      node.type.name === "field" &&
      (node.attrs["fieldKind"] === "simple" ||
        node.attrs["fieldKind"] === "complex") &&
      node.attrs["fieldResultIsFallback"] == null;
    // Source code runs are an editor round-trip carrier. The instruction is
    // the field's semantic identity; tracked simple fields become complex
    // fields with newly authored code runs on save.
    return Object.fromEntries(
      Object.entries(node.attrs).filter(
        ([name]) =>
          name !== "_docxFieldCode" &&
          (!equivalentFieldForms || name !== "fieldKind"),
      ),
    );
  }
  if (node.type.name === "image") {
    return Object.fromEntries(
      Object.entries(node.attrs)
        .filter(([name]) => name !== "docPrId")
        .map(([name, value]) => [
          name,
          name === "_docxRawXml" && typeof value === "string"
            ? drawingXmlWithoutIdentity(value)
            : value,
        ]),
    );
  }
  const local = CONVERSION_LOCAL_ATTRS[node.type.name];
  if (!local) return node.attrs;
  return Object.fromEntries(
    Object.entries(node.attrs).filter(([name]) => !local.includes(name)),
  );
};

const atomKey = (node: PMNode): string =>
  canonicalJson({
    type: node.type.name,
    attrs: documentFactAttrs(node),
    content: node.content.toJSON(),
  });

const atomBlockOf = (
  { node, from }: TextBlock,
  fieldResults: BuildCleanBlockTextOptions["fieldResults"],
  noteReferences: NoteReferenceReading,
): AtomBlock => {
  const clean = buildCleanBlockText(node, from, {
    fieldResults,
    noteReferences,
  });
  const references: NoteReferenceAt[] = [];
  for (const boundary of clean.structuralBoundaries) {
    if (boundary.type !== "noteReference") continue;
    references.push({
      noteType: boundary.noteType,
      noteId: boundary.noteId,
      offset: boundary.offset,
      length: boundary.length,
      from: boundary.from,
    });
  }
  const supported: InlineAtom[] = [];
  const unsupportedTopology: { offset: number; key: string }[] = [];
  node.descendants((child, relativePosition) => {
    if (child.isText) return false;
    if (!child.isInline || !child.isAtom) return true;
    const position = from + 1 + relativePosition;
    const offset = offsetAt({ offsets: clean.offsets, position });
    if (!isSupportedAtom(child)) {
      // The layout-only projection is derived from pageBreakRun. Its presence
      // cannot block reconciliation of the serializable carrier itself.
      if (child.type.name === "renderedPageBreak") return false;
      unsupportedTopology.push({
        offset,
        key: `${child.type.name}:${canonicalJson(documentFactAttrs(child))}`,
      });
      return false;
    }
    const prepared = prepareTargetInlineAtom(child);
    if (!prepared) {
      // An atom this comparison cannot detach from its package — an image with
      // no embedded media, a drawing that is a chart rather than a picture — is
      // one it cannot restore, not one that makes the story unalignable. Its
      // identity still has to match, so it joins the topology both sides are
      // compared on instead of abandoning the alignment for the whole story.
      unsupportedTopology.push({
        offset,
        key: `${child.type.name}:${canonicalJson(documentFactAttrs(child))}`,
      });
      return false;
    }
    supported.push({
      node: prepared,
      from: position,
      offset,
      key: atomKey(prepared),
      resultText: runFormattingInlineAtomResultText(prepared),
    });
    return false;
  });
  return {
    node,
    from,
    blockType: node.type.name,
    cleanText: clean.text,
    offsets: clean.offsets,
    supported,
    unsupportedTopology,
    noteReferences: references,
  };
};

export const atomBlocksOf = (
  doc: PMNode,
  fieldResults: BuildCleanBlockTextOptions["fieldResults"],
  noteReferences: NoteReferenceReading = "sourceText",
): AtomBlock[] =>
  textBlocksOf(doc).map((block) =>
    atomBlockOf(block, fieldResults, noteReferences),
  );

/** Compare supported inline atom identity after accept or reject projection. */
export const sameInlineAtoms = (
  leftDocument: PMNode,
  rightDocument: PMNode,
): boolean => {
  if (!hasSupportedAtom(leftDocument) && !hasSupportedAtom(rightDocument))
    return true;
  const leftBlocks = atomBlocksOf(leftDocument, "omitted");
  const rightBlocks = atomBlocksOf(rightDocument, "omitted");
  if (leftBlocks.length !== rightBlocks.length) return false;
  return leftBlocks.every((left, index) => {
    const right = rightBlocks[index];
    if (left.supported.length === 0 && right?.supported.length === 0)
      return true;
    return (
      right !== undefined &&
      left.cleanText === right.cleanText &&
      sameSupportedAtoms(left, right)
    );
  });
};
