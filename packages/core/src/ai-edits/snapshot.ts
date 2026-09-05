import type { Mark, Node as PMNode } from "prosemirror-model";

import { deriveBlockId } from "../types/block-id";
import { buildCleanBlockText } from "./clean-text";
import type {
  FolioAIBlock,
  FolioAIBlockAnchor,
  FolioAIBlockKind,
  FolioAIBlockPreviewRun,
  FolioAIBlockTableLocation,
  FolioAIEditSnapshot,
  FolioAITextRangeHandle,
} from "./types";

export const normalizeFolioAIBlockText = (text: string): string =>
  text.replace(/\s+/gu, " ").trim();

export const hashFolioAIBlockText = (text: string): string => {
  let hash = 5381;
  for (const character of text) {
    hash = (hash * 33 + (character.codePointAt(0) ?? 0)) % 2_147_483_647;
  }
  return `h${hash.toString(36)}`;
};

type CreateFolioAITextRangeHandleOptions = {
  blockId: string;
  text: string;
  startOffset: number;
  endOffset: number;
};

export const createFolioAITextRangeHandle = ({
  blockId,
  text,
  startOffset,
  endOffset,
}: CreateFolioAITextRangeHandleOptions): FolioAITextRangeHandle | null => {
  if (
    blockId.length === 0 ||
    !Number.isInteger(startOffset) ||
    !Number.isInteger(endOffset) ||
    startOffset < 0 ||
    endOffset <= startOffset ||
    endOffset > text.length
  ) {
    return null;
  }
  return {
    type: "textRange",
    story: "main",
    blockId,
    startOffset,
    endOffset,
    selectedTextHash: hashFolioAIBlockText(text.slice(startOffset, endOffset)),
  };
};

const TABLE_ROW_NODE_NAME = "tableRow";

const TABLE_ROLE_TABLE = "table";
const TABLE_ROLE_ROW = "row";

/**
 * One node on the path from the document to the block being visited: the node
 * itself, where it ends, and its index in its own parent.
 *
 * The walk below keeps this path instead of calling `doc.resolve(pos)` per
 * block. `resolve` re-descends from the root and finds each level's child by
 * scanning that level's fragment from index 0, so on a flat document of n
 * paragraphs it costs O(n) per block and the snapshot costs O(n^2). The path
 * is already known: a depth-first walk in document order visits every ancestor
 * before the block, so carrying it costs nothing and the snapshot is linear.
 */
type AncestorPathEntry = { node: PMNode; start: number; end: number; index: number };

/**
 * True for a `tableRow` marked `hidden` (OOXML `w:trPr/w:hidden` — Word never
 * renders the row). Nothing inside such a row may surface its text to the
 * AI/agent snapshot: an attacker DOCX could otherwise smuggle prompt-injection
 * instructions into a row that no human ever sees. The walk skips the row's
 * whole subtree, so a table nested inside a hidden row is hidden with it.
 */
const isHiddenTableRow = (node: PMNode): boolean =>
  node.type.name === TABLE_ROW_NODE_NAME && node.attrs["hidden"] === true;

type TableLocationOptions = {
  path: readonly AncestorPathEntry[];
  /** The visited block's index in its own parent. */
  blockIndex: number;
  tableIndexByStart: ReadonlyMap<number, number>;
};

/**
 * Locate the block inside its innermost enclosing table, or `undefined` when
 * it sits outside every table.
 */
const getTableLocation = ({
  path,
  blockIndex,
  tableIndexByStart,
}: TableLocationOptions): FolioAIBlockTableLocation | undefined => {
  for (let cellDepth = path.length - 1; cellDepth > 1; cellDepth--) {
    const cell = path[cellDepth];
    if (!cell) {
      continue;
    }
    const role = cell.node.type.spec["tableRole"];
    if (role !== "cell" && role !== "header_cell") {
      continue;
    }
    const row = path[cellDepth - 1];
    const table = path[cellDepth - 2];
    if (
      !row ||
      !table ||
      row.node.type.spec["tableRole"] !== TABLE_ROLE_ROW ||
      table.node.type.spec["tableRole"] !== TABLE_ROLE_TABLE
    ) {
      return undefined;
    }
    const tableIndex = tableIndexByStart.get(table.start);
    if (tableIndex === undefined) {
      return undefined;
    }
    return {
      tableIndex,
      rowIndex: row.index,
      cellIndex: cell.index,
      paragraphIndex: blockIndex,
    };
  }
  return undefined;
};

export const createFolioAIEditSnapshot = (doc: PMNode): FolioAIEditSnapshot => {
  const draftBlocks: {
    block: FolioAIBlock;
    anchor: Omit<FolioAIBlockAnchor, "hashOccurrenceCount">;
  }[] = [];
  const hashCounts = new Map<string, number>();
  const usedBlockIds = new Set<string>();
  // Table start position -> document-order index, filled by the walk below.
  // `descendants` reaches a table before any textblock inside it, so a nested
  // block's lookup always finds its table already numbered.
  const tableIndexByStart = new Map<number, number>();
  const emptyAnchorState: {
    candidate: { from: number; to: number; paraId: string | null } | null;
    textblockCount: number;
  } = { candidate: null, textblockCount: 0 };

  // The containers enclosing the node being visited, innermost last. Kept in
  // step with the walk so no block has to resolve its own position.
  const path: AncestorPathEntry[] = [];

  let blockIndex = 0;
  doc.descendants((node, pos, parent, index) => {
    while (pos >= (path.at(-1)?.end ?? Number.POSITIVE_INFINITY)) {
      path.pop();
    }
    if (!node.isTextblock) {
      if (isHiddenTableRow(node)) {
        return false;
      }
      if (node.type.spec["tableRole"] === TABLE_ROLE_TABLE) {
        tableIndexByStart.set(pos, tableIndexByStart.size);
      }
      if (!node.isLeaf) {
        path.push({ node, start: pos, end: pos + node.nodeSize, index });
      }
      return true;
    }

    // Snapshot the AI-facing text in its post-tracked-changes
    // form: existing deletion-marked runs are skipped, existing
    // insertion-marked runs are included as plain text. The model
    // would otherwise see "shallmust" smashed together in a block
    // mid-edit and write find/replace operations against that
    // confused string. Apply uses the same clean view to resolve
    // operation positions, so the offsets stay consistent.
    const { text } = buildCleanBlockText(node, pos);
    const normalizedText = normalizeFolioAIBlockText(text);
    if (normalizedText.length === 0) {
      if (parent?.type !== doc.type) {
        return true;
      }
      emptyAnchorState.textblockCount++;
      if (emptyAnchorState.candidate === null) {
        const paraIdAttr: unknown = node.attrs["paraId"];
        emptyAnchorState.candidate = {
          from: pos,
          to: pos + node.nodeSize,
          paraId: typeof paraIdAttr === "string" && paraIdAttr.length > 0 ? paraIdAttr : null,
        };
      }
      return true;
    }

    const textHash = hashFolioAIBlockText(normalizedText);
    hashCounts.set(textHash, (hashCounts.get(textHash) ?? 0) + 1);

    // Use the paragraph's Word `w14:paraId` (allocated by
    // `ParaIdAllocatorExtension` if the parsed DOCX didn't have one)
    // as the canonical block id everywhere: AI prompts, chip hrefs
    // (`#folio:<paraId>`), apply-tool blockIds, scrollToBlock. ParaIds
    // are stable across structural edits — no more "this chip points
    // at the wrong paragraph after an insertion-above" surprise.
    //
    // Shared with `apps/api/.../docx-blocks.ts` via `deriveBlockId`,
    // so a server-emitted citation id is always one of the two shapes
    // this snapshot produces (paraId verbatim or `seq-NNNN`).
    blockIndex++;
    const paraIdAttr: unknown = node.attrs["paraId"];
    const paraId = typeof paraIdAttr === "string" && paraIdAttr.length > 0 ? paraIdAttr : null;
    const id = deriveBlockId({
      paraId,
      index: blockIndex,
      taken: usedBlockIds,
    });
    usedBlockIds.add(id);
    const headingLevel = getHeadingLevel(node);
    const kind = getBlockKind(node, headingLevel);
    const displayLabel = getDisplayLabel(node);
    const styleId = getStyleId(node);
    const previewRuns = getPreviewRuns(node);
    const table = getTableLocation({ path, blockIndex: index, tableIndexByStart });

    draftBlocks.push({
      block: {
        id,
        kind,
        text,
        ...(headingLevel !== undefined && { headingLevel }),
        ...(displayLabel !== undefined && { displayLabel }),
        ...(styleId !== undefined && { styleId }),
        ...(previewRuns !== undefined && { previewRuns }),
        ...(table !== undefined && { table }),
      },
      anchor: {
        id,
        from: pos,
        to: pos + node.nodeSize,
        text,
        normalizedText,
        textHash,
      },
    });
    return true;
  });

  const blocks: FolioAIBlock[] = [];
  const anchors: Record<string, FolioAIBlockAnchor> = {};
  for (const draft of draftBlocks) {
    blocks.push(draft.block);
    anchors[draft.block.id] = {
      ...draft.anchor,
      hashOccurrenceCount: hashCounts.get(draft.anchor.textHash) ?? 0,
    };
  }

  const emptyAnchorCandidate = emptyAnchorState.candidate;
  if (blocks.length > 0 || emptyAnchorCandidate === null) {
    return { blocks, anchors };
  }

  const emptyDocumentAnchorId = deriveBlockId({
    paraId: emptyAnchorCandidate.paraId,
    index: 1,
    taken: usedBlockIds,
  });
  const normalizedText = "";
  anchors[emptyDocumentAnchorId] = {
    id: emptyDocumentAnchorId,
    from: emptyAnchorCandidate.from,
    to: emptyAnchorCandidate.to,
    text: "",
    normalizedText,
    textHash: hashFolioAIBlockText(normalizedText),
    hashOccurrenceCount: emptyAnchorState.textblockCount,
  };
  return { blocks, anchors, emptyDocumentAnchorId };
};

const getBlockKind = (node: PMNode, headingLevel: number | undefined): FolioAIBlockKind => {
  const listMarker: unknown = node.attrs["listMarker"];
  const numPr: unknown = node.attrs["numPr"];
  if (
    (typeof listMarker === "string" && listMarker.trim().length > 0) ||
    (numPr !== undefined && numPr !== null)
  ) {
    return "listItem";
  }

  if (headingLevel !== undefined) {
    return "heading";
  }

  return "paragraph";
};

const getHeadingLevel = (node: PMNode): number | undefined => {
  const outlineLevel: unknown = node.attrs["outlineLevel"];
  if (
    typeof outlineLevel === "number" &&
    Number.isInteger(outlineLevel) &&
    outlineLevel >= 0 &&
    outlineLevel <= 8
  ) {
    return outlineLevel + 1;
  }

  const styleId: unknown = node.attrs["styleId"];
  if (typeof styleId !== "string") {
    return undefined;
  }
  const match = /^heading(?<level>[1-9])$/iu.exec(styleId);
  const level = match?.groups?.["level"];
  return level === undefined ? undefined : Number.parseInt(level, 10);
};

const getDisplayLabel = (node: PMNode): string | undefined => {
  const listMarker: unknown = node.attrs["listMarker"];
  if (typeof listMarker === "string" && listMarker.trim().length > 0) {
    return listMarker.trim();
  }

  const styleId: unknown = node.attrs["styleId"];
  if (typeof styleId === "string" && /^heading/iu.test(styleId)) {
    return styleId;
  }

  return undefined;
};

const getStyleId = (node: PMNode): string | undefined => {
  const styleId: unknown = node.attrs["styleId"];
  return typeof styleId === "string" && styleId.length > 0 ? styleId : undefined;
};

type PreviewRunStyle = Omit<FolioAIBlockPreviewRun, "text">;

const DELETION_MARK = "deletion";

const getPreviewRuns = (node: PMNode): FolioAIBlockPreviewRun[] | undefined => {
  const runs: FolioAIBlockPreviewRun[] = [];
  const defaultStyle = getDefaultPreviewRunStyle(node);

  node.descendants((child) => {
    if (!child.isText || child.text === undefined) {
      return true;
    }
    if (child.marks.some((mark) => mark.type.name === DELETION_MARK)) {
      return false;
    }

    const style = getPreviewRunStyle(child.marks, defaultStyle);
    const previous = runs.at(-1);
    if (previous && samePreviewRunStyle(previous, style)) {
      previous.text += child.text;
      return false;
    }

    runs.push({ text: child.text, ...style });
    return false;
  });

  if (runs.every(isUnstyledPreviewRun)) {
    return undefined;
  }

  return runs;
};

const getDefaultPreviewRunStyle = (node: PMNode): PreviewRunStyle => {
  const formatting: unknown = node.attrs["defaultTextFormatting"];
  if (typeof formatting !== "object" || formatting === null) {
    return {};
  }

  return {
    ...getBooleanTextFormatting(formatting),
    ...getFontSizeTextFormatting(formatting),
    ...getFontFamilyTextFormatting(formatting),
    ...getColorTextFormatting(formatting),
  };
};

const getPreviewRunStyle = (
  marks: readonly Mark[],
  defaultStyle: PreviewRunStyle,
): PreviewRunStyle => {
  const style: PreviewRunStyle = { ...defaultStyle };

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        style.bold = true;
        break;
      case "italic":
        style.italic = true;
        break;
      case "underline":
        if (isUnderlineEnabled(mark.attrs["style"])) {
          style.underline = true;
        }
        break;
      case "strike":
        style.strike = true;
        break;
      case "fontSize": {
        const size = Number(mark.attrs["size"]);
        if (Number.isFinite(size) && size > 0) {
          style.fontSizePt = size / 2;
        }
        break;
      }
      case "fontFamily": {
        const fontFamily = getFontFamilyFromAttrs(mark.attrs);
        if (fontFamily !== undefined) {
          style.fontFamily = fontFamily;
        }
        break;
      }
      case "textColor": {
        const color = getColorFromAttrs(mark.attrs);
        if (color !== undefined) {
          style.color = color;
        }
        break;
      }
      default:
        break;
    }
  }

  return style;
};

const getBooleanTextFormatting = (formatting: object): PreviewRunStyle => ({
  ...(Reflect.get(formatting, "bold") === true && { bold: true }),
  ...(Reflect.get(formatting, "italic") === true && { italic: true }),
  ...(isUnderlineEnabled(Reflect.get(formatting, "underline")) && {
    underline: true,
  }),
  ...(Reflect.get(formatting, "strike") === true && { strike: true }),
});

const isUnderlineEnabled = (underline: unknown): boolean => {
  if (underline === undefined || underline === null || underline === false) {
    return false;
  }
  if (underline === true) {
    return true;
  }
  if (typeof underline === "string") {
    return underline !== "none";
  }
  if (typeof underline !== "object") {
    return false;
  }

  const style: unknown = Reflect.get(underline, "style");
  return style !== "none";
};

const getFontSizeTextFormatting = (formatting: object): PreviewRunStyle => {
  const fontSize = Number(Reflect.get(formatting, "fontSize"));
  if (!Number.isFinite(fontSize) || fontSize <= 0) {
    return {};
  }
  return { fontSizePt: fontSize / 2 };
};

const getFontFamilyTextFormatting = (formatting: object): PreviewRunStyle => {
  const fontFamilyValue: unknown = Reflect.get(formatting, "fontFamily");
  if (typeof fontFamilyValue !== "object" || fontFamilyValue === null) {
    return {};
  }

  const fontFamily = getFontFamilyFromAttrs(fontFamilyValue);
  return fontFamily === undefined ? {} : { fontFamily };
};

const getColorTextFormatting = (formatting: object): PreviewRunStyle => {
  const colorValue: unknown = Reflect.get(formatting, "color");
  if (typeof colorValue !== "object" || colorValue === null) {
    return {};
  }

  const color = getColorFromAttrs(colorValue);
  return color === undefined ? {} : { color };
};

const getFontFamilyFromAttrs = (attrs: object): string | undefined => {
  const ascii: unknown = Reflect.get(attrs, "ascii");
  if (typeof ascii === "string" && ascii.length > 0) {
    return ascii;
  }

  const hAnsi: unknown = Reflect.get(attrs, "hAnsi");
  if (typeof hAnsi === "string" && hAnsi.length > 0) {
    return hAnsi;
  }

  return undefined;
};

const getColorFromAttrs = (attrs: object): string | undefined => {
  const rgb: unknown = Reflect.get(attrs, "rgb") ?? Reflect.get(attrs, "val");
  if (typeof rgb !== "string" || !/^[0-9a-fA-F]{6}$/u.test(rgb)) {
    return undefined;
  }

  return `#${rgb}`;
};

const samePreviewRunStyle = (run: FolioAIBlockPreviewRun, style: PreviewRunStyle): boolean =>
  run.bold === style.bold &&
  run.italic === style.italic &&
  run.underline === style.underline &&
  run.strike === style.strike &&
  run.fontFamily === style.fontFamily &&
  run.fontSizePt === style.fontSizePt &&
  run.color === style.color;

const isUnstyledPreviewRun = ({
  bold,
  italic,
  underline,
  strike,
  fontFamily,
  fontSizePt,
  color,
}: FolioAIBlockPreviewRun): boolean =>
  bold === undefined &&
  italic === undefined &&
  underline === undefined &&
  strike === undefined &&
  fontFamily === undefined &&
  fontSizePt === undefined &&
  color === undefined;
