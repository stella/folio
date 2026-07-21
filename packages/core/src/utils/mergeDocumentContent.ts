/**
 * Merge Document Content Utility
 *
 * Composes two documents by appending one document's body content onto
 * another's, renumbering any list numbering the source carries so it can
 * never collide with numbering the target already defines.
 */

import type {
  AbstractNumbering,
  BlockContent,
  Document,
  ListRendering,
  NumberingDefinitions,
  NumberingInstance,
  Paragraph,
  Table,
} from "../types/document";

/**
 * Append `source`'s body content onto `target`'s. Both documents keep their
 * own identity — this returns a new `Document`; neither argument is mutated.
 *
 * `target` wins on everything except content and numbering: its style set,
 * theme, font table, settings, and section/page geometry are unchanged, so
 * merging markdown (or any other document) into a styled preset keeps that
 * preset's look. Only `source`'s numbering — the `w:abstractNum`/`w:num`
 * definitions its content's `numPr` references — travels with it, and every
 * `abstractNumId`/`numId` it carries is renumbered to sit strictly above
 * whatever `target` already uses. This is unconditional (not just on a
 * detected collision): it is simpler to reason about, and makes repeated
 * merges into the same target safe too (merging two sources that each mint
 * `numId` 1 does not collide with each other either).
 *
 * A typical caller composes `fromMarkdown` with a styled preset:
 *
 * ```ts
 * const target = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
 * const merged = mergeDocumentContent(target, fromMarkdown(markdown));
 * const bytes = await createDocx(merged);
 * ```
 */
export function mergeDocumentContent(target: Document, source: Document): Document {
  const targetNumbering = target.package.numbering;
  const sourceNumbering = source.package.numbering;

  if (
    sourceNumbering === undefined ||
    (sourceNumbering.abstractNums.length === 0 && sourceNumbering.nums.length === 0)
  ) {
    return {
      ...target,
      package: {
        ...target.package,
        document: {
          ...target.package.document,
          content: [...target.package.document.content, ...source.package.document.content],
        },
      },
    };
  }

  const remappedSource = remapContentNumbering(source.package.document.content, sourceNumbering);

  return {
    ...target,
    package: {
      ...target.package,
      document: {
        ...target.package.document,
        content: [...target.package.document.content, ...remappedSource.content],
      },
      numbering: {
        abstractNums: [
          ...(targetNumbering?.abstractNums ?? []),
          ...remappedSource.numbering.abstractNums,
        ],
        nums: [...(targetNumbering?.nums ?? []), ...remappedSource.numbering.nums],
      },
    },
  };

  function remapContentNumbering(
    content: BlockContent[],
    numbering: NumberingDefinitions,
  ): { content: BlockContent[]; numbering: NumberingDefinitions } {
    const abstractNumIdBase = Math.max(
      0,
      ...(targetNumbering?.abstractNums ?? []).map((abstractNum) => abstractNum.abstractNumId),
    );
    const numIdBase = Math.max(0, ...(targetNumbering?.nums ?? []).map((num) => num.numId));

    const abstractNumIdRemap = new Map<number, number>();
    const remappedAbstractNums: AbstractNumbering[] = numbering.abstractNums.map(
      (abstractNum, index) => {
        const remappedAbstractNumId = abstractNumIdBase + index + 1;
        abstractNumIdRemap.set(abstractNum.abstractNumId, remappedAbstractNumId);
        return { ...abstractNum, abstractNumId: remappedAbstractNumId };
      },
    );

    const numIdRemap = new Map<number, number>();
    const remappedNums: NumberingInstance[] = numbering.nums.map((num, index) => {
      const remappedNumId = numIdBase + index + 1;
      numIdRemap.set(num.numId, remappedNumId);
      return {
        ...num,
        numId: remappedNumId,
        abstractNumId: abstractNumIdRemap.get(num.abstractNumId) ?? num.abstractNumId,
      };
    });

    return {
      content: content.map((block) => remapBlock(block, numIdRemap, abstractNumIdRemap)),
      numbering: { abstractNums: remappedAbstractNums, nums: remappedNums },
    };
  }
}

function remapBlock(
  block: BlockContent,
  numIdRemap: Map<number, number>,
  abstractNumIdRemap: Map<number, number>,
): BlockContent {
  if (block.type === "paragraph") {
    return remapParagraph(block, numIdRemap, abstractNumIdRemap);
  }
  if (block.type === "table") {
    return remapTable(block, numIdRemap, abstractNumIdRemap);
  }
  // blockSdt — content controls can nest paragraphs/tables/other controls.
  // `BlockContent` is meant to be exhaustive here, but this helper also runs
  // on content from hand-built `Document` inputs (not necessarily produced by
  // `parseDocx`), so guard the shape at runtime rather than trusting the
  // static type: a block whose `type` matches neither "paragraph" nor
  // "table" but that also carries no `content` array passes through
  // unchanged instead of throwing.
  if ("content" in block && Array.isArray(block.content)) {
    return {
      ...block,
      content: block.content.map((child) => remapBlock(child, numIdRemap, abstractNumIdRemap)),
    };
  }
  return block;
}

function remapTable(
  table: Table,
  numIdRemap: Map<number, number>,
  abstractNumIdRemap: Map<number, number>,
): Table {
  return {
    ...table,
    rows: table.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        content: cell.content.map((item) => remapCellItem(item, numIdRemap, abstractNumIdRemap)),
      })),
    })),
  };
}

/**
 * `TableCell.content` is typed as `(Paragraph | Table)[]` — narrower than
 * `BlockContent` — because folio's own DOCX parser flattens a `w:sdt` inside
 * a cell into its `sdtContent` children (see `tableParser.ts`) rather than
 * keeping a `blockSdt` wrapper. A hand-built `Document` is not bound by that
 * invariant, so delegate to `remapBlock` (which fully handles paragraph,
 * table, and container blocks, including nested tables) and only accept its
 * result back into the cell when it is still a `Paragraph | Table` — an
 * anomalous `blockSdt` (or any other shape `remapBlock` had to pass through
 * unchanged) is left as-is rather than smuggled into a field the type system
 * says can't hold it.
 */
function remapCellItem(
  item: Paragraph | Table,
  numIdRemap: Map<number, number>,
  abstractNumIdRemap: Map<number, number>,
): Paragraph | Table {
  const remapped = remapBlock(item, numIdRemap, abstractNumIdRemap);
  return remapped.type === "blockSdt" ? item : remapped;
}

function remapParagraph(
  paragraph: Paragraph,
  numIdRemap: Map<number, number>,
  abstractNumIdRemap: Map<number, number>,
): Paragraph {
  const numId = paragraph.formatting?.numPr?.numId;
  const remappedNumId = numId === undefined ? undefined : numIdRemap.get(numId);
  if (remappedNumId === undefined) {
    return paragraph;
  }

  return {
    ...paragraph,
    formatting: {
      ...paragraph.formatting,
      numPr: { ...paragraph.formatting?.numPr, numId: remappedNumId },
    },
    ...(paragraph.listRendering && {
      listRendering: remapListRendering(paragraph.listRendering, remappedNumId, abstractNumIdRemap),
    }),
  };
}

function remapListRendering(
  rendering: ListRendering,
  remappedNumId: number,
  abstractNumIdRemap: Map<number, number>,
): ListRendering {
  return {
    ...rendering,
    numId: remappedNumId,
    ...(rendering.abstractNumId !== undefined && {
      abstractNumId: abstractNumIdRemap.get(rendering.abstractNumId) ?? rendering.abstractNumId,
    }),
  };
}
