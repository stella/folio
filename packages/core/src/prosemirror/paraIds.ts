/**
 * Paragraph-id allocation shared by load-time seeding and the live
 * allocator plugin (`ParaIdAllocatorExtension`).
 *
 * Split out of
 * https://github.com/eigenpal/docx-editor/blob/main/packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts
 * (Apache-2.0); see that module for the id lifecycle rules.
 */
import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { EditorState, Plugin, PluginKey } from "prosemirror-state";

import {
  recreateProseNodeWithParagraphPropertySource,
  setProseParagraphMarkupWithPropertySource,
  transferProseParagraphPropertySource,
} from "../docx/paragraphPropertySource";
import { prosePositionalParagraphId } from "../docx/unnamedParagraphIds";
import { deterministicHexId } from "../utils/hexId";
import { ignoreTrackedChanges } from "./paragraphChangeTrackerMeta";
import { createPluginKeyOwner } from "./pluginKeyOwner";

export type ParagraphPropertySourceSeed = {
  contract: string | null;
  tokens: ReadonlySet<string>;
};

export const paraIdAllocatorKey: Readonly<{
  key: string;
  createKey: () => PluginKey<ParagraphPropertySourceSeed>;
  owns: (plugin: Plugin<ParagraphPropertySourceSeed>) => boolean;
  get: (state: EditorState) => Plugin<ParagraphPropertySourceSeed> | undefined;
  getState: (state: EditorState) => ParagraphPropertySourceSeed | undefined;
}> = createPluginKeyOwner<ParagraphPropertySourceSeed>("paraIdAllocator");

export const isUsableParaId = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0 && value !== "00000000";

/** Omitted grid cells render absent positions and own no package paragraphs. */
const isOmittedGridCell = (node: PMNode): boolean =>
  (node.type.spec["tableRole"] === "cell" ||
    node.type.spec["tableRole"] === "header_cell") &&
  (node.attrs["_omittedGridSlot"] === "before" ||
    node.attrs["_omittedGridSlot"] === "after");

export const visitAuthoredParagraphs = (
  doc: PMNode,
  visit: (node: PMNode, pos: number) => void,
): void => {
  doc.descendants((node, pos) => {
    if (isOmittedGridCell(node)) return false;
    if (node.type.name !== "paragraph") return true;
    visit(node, pos);
    return false;
  });
};

/**
 * The ids a loaded document's paragraphs take, in document order: `null`
 * keeps the paragraph's own id, a string replaces it.
 *
 * This is the only load-time allocator. The live editor and the headless
 * reviewer both seed through it, so the same document gives the same block ids
 * on either side: an id read from the editor resolves when the edit is applied
 * headlessly, and re-parsing the same bytes mints the same ids. A paragraph
 * the package names no id for takes the id its place in the package gives it
 * (the one `ensureParaIds` writes, so a canonical session agrees); one the
 * parser could not place, such as a paragraph a host built, takes an id
 * derived from its text and its document ordinal, never drawn at random.
 * Authored ids are reserved first, so a minted id never takes one a later
 * paragraph owns; of a duplicated authored id, the first occurrence keeps it.
 */
const planLoadParaIds = (doc: PMNode): (string | null)[] => {
  const paragraphs: PMNode[] = [];
  const taken = new Set<string>();
  visitAuthoredParagraphs(doc, (node) => {
    paragraphs.push(node);
    const id = node.attrs["paraId"];
    if (isUsableParaId(id)) {
      taken.add(id);
      return;
    }
    const positional = prosePositionalParagraphId(node);
    if (positional !== undefined) taken.add(positional);
  });

  const kept = new Set<string>();
  return paragraphs.map((paragraph, index) => {
    const id = paragraph.attrs["paraId"];
    if (isUsableParaId(id) && !kept.has(id)) {
      kept.add(id);
      return null;
    }
    // The id the package's position gives it, as a canonical session reads
    // the package, unless an authored id already holds it.
    const positional = prosePositionalParagraphId(paragraph);
    if (positional !== undefined && !kept.has(positional)) {
      kept.add(positional);
      return positional;
    }
    const seed = `${paragraph.textContent}:${String(index + 1)}`;
    let minted = deterministicHexId(seed);
    for (let salt = 1; taken.has(minted); salt++) {
      minted = deterministicHexId(`${seed}:${String(salt)}`);
    }
    taken.add(minted);
    return minted;
  });
};

/** A paragraph's attrs once it holds an id folio minted rather than read. */
const mintedParagraphAttrs = (
  paragraph: PMNode,
  paraId: string,
): Record<string, unknown> => ({
  ...paragraph.attrs,
  paraId,
  // Snapshots report a minted id as positional: no package names it.
  ...(paragraph.type.spec.attrs?.["idStability"] !== undefined && {
    idStability: "positional",
  }),
});

/**
 * Allocate initial paragraph IDs before constructing editor state.
 *
 * Rebuilding only changed branches avoids applying one ProseMirror transaction
 * step per missing paragraph. Live edits still use the transaction-based plugin
 * below so history, mappings, and duplicate ownership retain their semantics.
 */
export const ensureParaIdsInDoc = (doc: PMNode): PMNode => {
  const plan = planLoadParaIds(doc);
  if (plan.every((paraId) => paraId === null)) {
    return doc;
  }

  let ordinal = 0;
  const rewrite = (parent: PMNode): Fragment => {
    let changed = false;
    const children: PMNode[] = [];
    parent.forEach((child) => {
      if (isOmittedGridCell(child)) {
        children.push(child);
        return;
      }
      let next = child;
      if (child.type.name === "paragraph") {
        const minted = plan[ordinal];
        ordinal += 1;
        if (minted !== null && minted !== undefined) {
          next = recreateProseNodeWithParagraphPropertySource(child, {
            attrs: mintedParagraphAttrs(child, minted),
          });
        }
        const paraId = next.attrs["paraId"];
        if (typeof paraId === "string") {
          transferProseParagraphPropertySource(next, child, paraId);
        }
      } else if (child.childCount > 0) {
        const content = rewrite(child);
        if (content !== child.content) {
          next = recreateProseNodeWithParagraphPropertySource(child, {
            content,
          });
        }
      }
      if (next !== child) {
        changed = true;
      }
      children.push(next);
    });
    return changed ? Fragment.fromArray(children) : parent.content;
  };

  return recreateProseNodeWithParagraphPropertySource(doc, {
    content: rewrite(doc),
  });
};

/** {@link ensureParaIdsInDoc} for a state that already exists. */
export const ensureParaIdsInState = (state: EditorState): EditorState => {
  const plan = planLoadParaIds(state.doc);
  if (plan.every((paraId) => paraId === null)) {
    return state;
  }

  const tr = state.tr;
  let ordinal = 0;
  // Markup-only steps keep every position, so the old doc's positions hold.
  visitAuthoredParagraphs(state.doc, (node, pos) => {
    const minted = plan[ordinal];
    ordinal += 1;
    if (minted !== null && minted !== undefined) {
      setProseParagraphMarkupWithPropertySource({
        attrs: mintedParagraphAttrs(node, minted),
        ownership: "transfer-allocated-id",
        pos,
        transaction: tr,
      });
    }
  });
  ignoreTrackedChanges(tr);
  tr.setMeta(paraIdAllocatorKey, "allocated");
  tr.setMeta("addToHistory", false);
  return state.apply(tr);
};
