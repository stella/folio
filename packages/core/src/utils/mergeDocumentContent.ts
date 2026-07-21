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
  return {
    ...block,
    content: block.content.map((child) => remapBlock(child, numIdRemap, abstractNumIdRemap)),
  };
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
        content: cell.content.map((item) =>
          item.type === "paragraph"
            ? remapParagraph(item, numIdRemap, abstractNumIdRemap)
            : remapTable(item, numIdRemap, abstractNumIdRemap),
        ),
      })),
    })),
  };
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
