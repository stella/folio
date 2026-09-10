import { panic } from "better-result";

import type {
  BlockContent,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  TableCell,
  TableRow,
  TrackedRunContent,
} from "../types/document";

export type PageBreakRunSourceDescendantIndex = {
  containsPageBreakRun: (content: readonly BlockContent[]) => boolean;
  paragraphFeatures: (paragraph: Paragraph) => PageBreakRunSourceParagraphFeatures;
};

export type PageBreakRunSourceParagraphFeatures = {
  hasPageBreakRun: boolean;
  hasTextBoxShape: boolean;
  pageBreakSharesTextBoxShape: boolean;
};

type ParagraphRunTreeNode = ParagraphContent | TrackedRunContent | InlineSdt["content"][number];

const SOURCE_FLAGS = {
  containsPageBreak: 1,
  paragraphPageBreak: 2,
  paragraphTextBoxShape: 4,
  pageBreakSharesTextBoxShape: 8,
} as const;

const hasFlag = (flags: number, flag: number): boolean => (flags & flag) !== 0;

/** Index authored page-break descendants once for one source conversion. */
export const buildPageBreakRunSourceDescendantIndex = (
  rootContent: readonly BlockContent[],
): PageBreakRunSourceDescendantIndex => {
  const flagsByObject = new WeakMap<object, number>();

  const inspectObject = (node: object, inspectChildren: () => number): number => {
    const cached = flagsByObject.get(node);
    if (cached !== undefined) {
      return cached;
    }

    // Mark the object before descending so malformed cyclic input remains bounded.
    flagsByObject.set(node, 0);
    const flags = inspectChildren();
    flagsByObject.set(node, flags);
    return flags;
  };

  const inspectMany = <T>(items: readonly T[], inspect: (item: T) => number): number => {
    let flags = 0;
    for (const item of items) {
      flags |= inspect(item);
    }
    return flags;
  };

  const inspectRunContent = (content: RunContent): number =>
    inspectObject(content, () => {
      switch (content.type) {
        case "break":
          return content.breakType === "page"
            ? SOURCE_FLAGS.containsPageBreak | SOURCE_FLAGS.paragraphPageBreak
            : 0;
        case "shape": {
          const { textBody } = content.shape;
          if (!textBody) {
            return 0;
          }
          const descendantFlags = inspectBlocks(textBody.content);
          return (
            (descendantFlags & SOURCE_FLAGS.containsPageBreak) | SOURCE_FLAGS.paragraphTextBoxShape
          );
        }
        case "drawing":
        case "endnoteRef":
        case "fieldChar":
        case "footnoteRef":
        case "instrText":
        case "noBreakHyphen":
        case "renderedPageBreak":
        case "softHyphen":
        case "symbol":
        case "tab":
        case "text":
          return 0;
        default: {
          const unsupported: never = content;
          return unsupported;
        }
      }
    });

  const inspectRun = (run: Run): number =>
    inspectObject(run, () => {
      const flags = inspectMany(run.content, inspectRunContent);
      if (
        hasFlag(flags, SOURCE_FLAGS.paragraphPageBreak) &&
        hasFlag(flags, SOURCE_FLAGS.paragraphTextBoxShape)
      ) {
        return flags | SOURCE_FLAGS.pageBreakSharesTextBoxShape;
      }
      return flags;
    });

  const inspectParagraphContent = (content: ParagraphRunTreeNode): number => {
    if (content.type === "run") {
      return inspectRun(content);
    }

    return inspectObject(content, () => {
      switch (content.type) {
        case "hyperlink":
          return inspectMany(content.children, inspectParagraphContent);
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
        case "inlineSdt":
          return inspectMany(content.content, inspectParagraphContent);
        case "simpleField": {
          const fieldResultFlags = inspectMany(content.content, inspectParagraphContent);
          if (
            hasFlag(fieldResultFlags, SOURCE_FLAGS.paragraphPageBreak) &&
            hasFlag(fieldResultFlags, SOURCE_FLAGS.paragraphTextBoxShape)
          ) {
            return fieldResultFlags | SOURCE_FLAGS.pageBreakSharesTextBoxShape;
          }
          return fieldResultFlags;
        }
        case "complexField": {
          const fieldCodeFlags = inspectMany(content.fieldCode, inspectRun);
          const fieldResultFlags = inspectMany(content.fieldResult, inspectRun);
          const fieldFlags = fieldCodeFlags | fieldResultFlags;
          if (
            hasFlag(fieldResultFlags, SOURCE_FLAGS.paragraphPageBreak) &&
            hasFlag(fieldResultFlags, SOURCE_FLAGS.paragraphTextBoxShape)
          ) {
            return fieldFlags | SOURCE_FLAGS.pageBreakSharesTextBoxShape;
          }
          return fieldFlags;
        }
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "commentReference":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
        case "mathEquation":
          return 0;
        default: {
          const unsupported: never = content;
          return unsupported;
        }
      }
    });
  };

  const inspectTableCell = (cell: TableCell): number =>
    inspectObject(cell, () => inspectBlocks(cell.content));

  const inspectTableRow = (row: TableRow): number =>
    inspectObject(row, () => inspectMany(row.cells, inspectTableCell));

  const inspectBlock = (block: BlockContent): number =>
    inspectObject(block, () => {
      switch (block.type) {
        case "paragraph": {
          const paragraphFlags = inspectMany(block.content, inspectParagraphContent);
          return paragraphFlags;
        }
        case "table":
          return inspectMany(block.rows, inspectTableRow);
        case "blockSdt":
          return inspectBlocks(block.content);
        default: {
          const unsupported: never = block;
          return unsupported;
        }
      }
    });

  function inspectBlocks(content: readonly BlockContent[]): number {
    const cached = flagsByObject.get(content);
    if (cached !== undefined) {
      return cached;
    }

    flagsByObject.set(content, 0);
    const flags = inspectMany(content, inspectBlock);
    flagsByObject.set(content, flags);
    return flags;
  }

  inspectBlocks(rootContent);

  return {
    containsPageBreakRun: (content) => {
      const flags = flagsByObject.get(content);
      if (flags === undefined) {
        panic("Page-break source ownership was queried outside its conversion index");
      }
      return hasFlag(flags, SOURCE_FLAGS.containsPageBreak);
    },
    paragraphFeatures: (paragraph) => {
      const flags = flagsByObject.get(paragraph);
      if (flags === undefined) {
        panic("Page-break paragraph ownership was queried outside its conversion index");
      }
      return {
        hasPageBreakRun: hasFlag(flags, SOURCE_FLAGS.paragraphPageBreak),
        hasTextBoxShape: hasFlag(flags, SOURCE_FLAGS.paragraphTextBoxShape),
        pageBreakSharesTextBoxShape: hasFlag(flags, SOURCE_FLAGS.pageBreakSharesTextBoxShape),
      };
    },
  };
};
