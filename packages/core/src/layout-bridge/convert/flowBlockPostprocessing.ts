/**
 * Post-conversion passes over the flow block list: terminal empty-paragraph
 * suppression, leading empty outline reservation, and run-in paragraph
 * merging.
 */

import { headingLevelOf } from "@stll/docx-core/model";
import type {
  FlowBlock,
  ParagraphBlock,
  ParagraphAttrs,
  SdtGroup,
} from "../../layout-engine/types";

/**
 * Word keeps a final empty body paragraph after a table as an editable anchor,
 * but that final anchor does not create a page of its own. Earlier authored
 * empty paragraphs retain their height and may carry the document onto a blank
 * page. Preserve every block and PM range while collapsing only the final one.
 */
function isPaintlessTerminalParagraph(block: FlowBlock | undefined): block is ParagraphBlock {
  if (block?.kind !== "paragraph" || block.runs.length !== 0) {
    return false;
  }

  const attrs = block.attrs;
  if (attrs?.suppressEmptyParagraphHeight === false) {
    return false;
  }
  return !(
    (attrs?.listMarker !== undefined && !attrs.listMarkerHidden) ||
    attrs?.borders?.top ||
    attrs?.borders?.bottom ||
    attrs?.borders?.left ||
    attrs?.borders?.right ||
    attrs?.borders?.between ||
    attrs?.borders?.bar ||
    attrs?.shading ||
    attrs?.spacingExplicit?.before ||
    attrs?.spacingExplicit?.after ||
    attrs?.pageBreakBefore ||
    attrs?.renderedPageBreakBefore
  );
}

export function suppressFinalEmptyParagraphAfterTable(blocks: FlowBlock[]): void {
  let suffixStart = blocks.length;
  while (suffixStart > 0 && isPaintlessTerminalParagraph(blocks[suffixStart - 1])) {
    suffixStart -= 1;
  }

  if (
    suffixStart === blocks.length ||
    suffixStart === 0 ||
    blocks[suffixStart - 1]?.kind !== "table"
  ) {
    return;
  }

  const finalBlock = blocks.at(-1);
  if (isPaintlessTerminalParagraph(finalBlock)) {
    finalBlock.attrs = { ...finalBlock.attrs, suppressEmptyParagraphHeight: true };
  }
}

export function suppressFinalParagraphInRepeatedEmptySuffix(blocks: FlowBlock[]): void {
  let suffixStart = blocks.length;
  while (suffixStart > 0 && isPaintlessTerminalParagraph(blocks[suffixStart - 1])) {
    suffixStart -= 1;
  }

  if (
    suffixStart === 0 ||
    blocks.length - suffixStart < 2 ||
    blocks[suffixStart - 1]?.kind === "table"
  ) {
    return;
  }

  const finalBlock = blocks.at(-1);
  if (isPaintlessTerminalParagraph(finalBlock)) {
    finalBlock.attrs = { ...finalBlock.attrs, suppressEmptyParagraphHeight: true };
  }
}

export function reserveLeadingEmptyOutlineHeight(blocks: FlowBlock[]): void {
  const firstBlock = blocks.at(0);
  if (
    firstBlock?.kind !== "paragraph" ||
    firstBlock.runs.length !== 0 ||
    headingLevelOf(firstBlock.attrs?.outlineLevel) !== 0
  ) {
    return;
  }

  firstBlock.attrs = { ...firstBlock.attrs, reserveEmptyOutlineHeight: true };
}

/**
 * Merge consecutive paragraph blocks where the first carries
 * `runInWithNext` (`<w:specVanish/>` on the paragraph mark).
 *
 * Word's run-in heading feature renders the next paragraph inline on
 * the same line, so for layout we collapse the pair into one
 * ParagraphBlock with combined runs. The merged block keeps the first
 * paragraph's attrs (heading formatting, list marker, indent) and
 * extends pmEnd to the second paragraph's range so click-to-position
 * resolution still maps both ranges back to body content.
 *
 * Chains: runInWithNext on the merged block is dropped because the
 * second paragraph's mark wasn't `specVanish`. If a chain of
 * specVanish paragraphs needs collapsing (rare in practice), the loop
 * naturally handles it by re-inspecting the merged block's flag (we
 * preserve runInWithNext only when the second paragraph itself has
 * specVanish).
 */
function sdtGroupStacksEqual(a: SdtGroup[] | undefined, b: SdtGroup[] | undefined): boolean {
  // pmPos is unique per SDT instance within a single toFlowBlocks call, so
  // comparing the pmPos stacks is enough to tell "same membership" vs.
  // "different membership" without copying the rest of the group payload.
  const lenA = a?.length ?? 0;
  const lenB = b?.length ?? 0;
  if (lenA !== lenB) {
    return false;
  }
  if (lenA === 0) {
    return true;
  }
  for (let i = 0; i < lenA; i++) {
    if (a?.[i]?.pmPos !== b?.[i]?.pmPos) {
      return false;
    }
  }
  return true;
}

export function mergeRunInParagraphs(blocks: FlowBlock[]): FlowBlock[] {
  const out: FlowBlock[] = [];
  for (let i = 0; i < blocks.length; i++) {
    let current = blocks[i];
    if (!current) {
      continue;
    }
    // Chain merge: keep folding consecutive paragraphs while the
    // *current* (possibly already-merged) block carries
    // `runInWithNext` and the next block is also a paragraph. Per
    // ECMA-376 §17.3.1.32 and Word's behaviour, a sequence of
    // `<w:specVanish/>` paragraphs flows inline through the first
    // body paragraph that lacks it (Codex PR #258 review).
    while (
      current.kind === "paragraph" &&
      (current as ParagraphBlock).attrs?.runInWithNext &&
      i + 1 < blocks.length
    ) {
      const next = blocks[i + 1];
      if (!next || next.kind !== "paragraph") {
        break;
      }
      const a = current as ParagraphBlock;
      const b = next as ParagraphBlock;
      // Stop merging when the two paragraphs sit in different SDT stacks.
      // The merged ParagraphBlock would inherit only `a`'s sdtGroups via
      // spread, so a `<w:specVanish/>` adjacent to a paragraph across an
      // SDT boundary would either claim outside text as part of the SDT
      // or strip SDT membership from inside text — chrome and widget
      // click targets would line up against the wrong content range.
      if (!sdtGroupStacksEqual(a.sdtGroups, b.sdtGroups)) {
        break;
      }
      const mergedAttrs: ParagraphAttrs = { ...a.attrs };
      // Heading typically has no spaceAfter; the body's spaceAfter
      // governs the merged paragraph's trailing gap.
      if (b.attrs?.spacing?.after !== undefined) {
        mergedAttrs.spacing = {
          ...mergedAttrs.spacing,
          after: b.attrs.spacing.after,
        };
      }
      // Carry forward `runInWithNext` only if the *consumed* second
      // paragraph itself was specVanish — the while condition above
      // then triggers another fold against the paragraph after it.
      if (b.attrs?.runInWithNext) {
        mergedAttrs.runInWithNext = true;
      } else {
        delete mergedAttrs.runInWithNext;
      }
      const merged: ParagraphBlock = {
        ...a,
        runs: [...a.runs, ...b.runs],
        attrs: mergedAttrs,
      };
      const mergedPmEnd = b.pmEnd ?? a.pmEnd;
      if (mergedPmEnd !== undefined) {
        merged.pmEnd = mergedPmEnd;
      }
      current = merged;
      i += 1; // consumed `next`; fold further if the merged block still has the flag
    }
    out.push(current);
  }
  return out;
}
