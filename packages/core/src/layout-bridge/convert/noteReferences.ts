/**
 * Footnote and endnote reference detection: collects note references from a
 * document and recognises preserved note reference marks.
 */

import type { FlowBlock } from "../../layout-engine/types";

const NOTE_REFERENCE_MARK_XML = /^<(?:[^\s:/>]+:)?(?:footnoteRef|endnoteRef)[\s/>]/u;

/**
 * Runs with `idKey` set, in document order: the note references a block list
 * makes, through table cells (and tables within cells) and text boxes.
 */
export function collectNoteRefs(
  blocks: readonly FlowBlock[],
  idKey: "footnoteRefId" | "endnoteRefId",
): { noteId: number; pmPos: number }[] {
  const refs: { noteId: number; pmPos: number }[] = [];

  const walk = (containerBlocks: readonly FlowBlock[]): void => {
    for (const block of containerBlocks) {
      if (block.kind === "paragraph") {
        for (const run of block.runs) {
          if (run.kind !== "text") {
            continue;
          }
          const noteId = run[idKey];
          if (noteId !== undefined) {
            refs.push({ noteId, pmPos: run.pmStart ?? 0 });
          }
        }
      } else if (block.kind === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walk(cell.blocks);
          }
        }
      } else if (block.kind === "textBox") {
        walk(block.content);
      }
    }
  };

  walk(blocks);
  return refs;
}

/** Whether captured run-child markup is a note story's `w:footnoteRef`/`w:endnoteRef`. */
export function isNoteReferenceMarkXml(xml: string): boolean {
  return NOTE_REFERENCE_MARK_XML.test(xml);
}
