/**
 * Sections inside table cells.
 *
 * A section ends at a paragraph outside every table, so no operation may end
 * one in a cell. A document may arrive holding one: an operation that leaves
 * it where it stands applies, and one after which a section stands anywhere
 * it did not stand before is refused, whether it makes, moves, copies or
 * restores it. The answer is read off the two documents alone: the same
 * operation on the same content has the same answer, however the document
 * came to be.
 */
import type { BlockContent, Document, Paragraph } from "../model/document";
import { type BlockListCheck, blockListCheck } from "./blockListCheck";
import { idKey } from "./ids";
import { documentStories, findStoryBody, paragraphTextBoxes } from "./stories";

const boxContents = (paragraph: Paragraph): readonly BlockContent[][] =>
  paragraphTextBoxes(paragraph).map(({ leaf }) => leaf.shape.textBody.content);

const blocksEndSection: BlockListCheck = blockListCheck({
  paragraph: (paragraph) =>
    paragraph.sectionProperties !== undefined ||
    boxContents(paragraph).some((blocks) => blocksEndSection(blocks)),
});

const cellsEndSection: BlockListCheck = blockListCheck({
  paragraph: (paragraph) =>
    boxContents(paragraph).some((blocks) => cellsEndSection(blocks)),
  cell: blocksEndSection,
});

/** Whether a paragraph anywhere in a table cell ends a section. */
export const hasSectionInTableCell = (document: Document): boolean =>
  documentStories(document).some((story) =>
    cellsEndSection(findStoryBody(document, story)?.content ?? []),
  );

/** A container, or paragraph, among the ones beside it, each by a paragraph of its own. */
type Level = { anchor: string | undefined; siblings: readonly string[] };
/** Where a paragraph stands: its containers from the story's blocks down, and itself. */
type Place = readonly Level[];

/** The blocks of a list, those of its content controls and custom XML in their place. */
const flatBlocks = (blocks: readonly BlockContent[]): BlockContent[] =>
  blocks.flatMap((block) =>
    block.type === "blockSdt" || block.type === "blockCustomXml"
      ? flatBlocks(block.content)
      : [block],
  );

const collectIds = (
  blocks: readonly BlockContent[],
  ids: Set<string>,
): Set<string> => {
  for (const block of flatBlocks(blocks)) {
    if (block.type === "table")
      for (const row of block.rows)
        for (const cell of row.cells) collectIds(cell.content, ids);
    if (block.type !== "paragraph") continue;
    if (block.paraId !== undefined) ids.add(idKey(block.paraId));
    for (const content of boxContents(block)) collectIds(content, ids);
  }
  return ids;
};

type Survivors = ReadonlySet<string>;

/** The first paragraph in `blocks` that stands in both documents. */
const firstSurvivor = (
  blocks: readonly BlockContent[],
  survivors: Survivors,
): string | undefined => {
  for (const block of flatBlocks(blocks)) {
    if (block.type === "table")
      for (const row of block.rows)
        for (const cell of row.cells) {
          const found = firstSurvivor(cell.content, survivors);
          if (found !== undefined) return found;
        }
    if (block.type !== "paragraph") continue;
    const id = block.paraId === undefined ? undefined : idKey(block.paraId);
    if (id !== undefined && survivors.has(id)) return id;
    for (const content of boxContents(block)) {
      const found = firstSurvivor(content, survivors);
      if (found !== undefined) return found;
    }
  }
  return undefined;
};

type SiblingOptions<Item> = {
  items: readonly Item[];
  blocksOf: (item: Item) => readonly BlockContent[];
  survivors: Survivors;
};

/** Each item with its level among the others. */
const levelled = <Item>({
  items,
  blocksOf,
  survivors,
}: SiblingOptions<Item>): { item: Item; level: Level }[] => {
  const anchors = items.map((item) => firstSurvivor(blocksOf(item), survivors));
  const siblings = anchors.filter((anchor) => anchor !== undefined);
  return items.map((item, index) => ({
    item,
    level: { anchor: anchors[index], siblings },
  }));
};

type PlacesOptions = {
  blocks: readonly BlockContent[];
  survivors: Survivors;
  above: Place;
  inCell: boolean;
  out: Map<string, Place>;
};

/** Where each section inside a table cell ends, by its paragraph. */
const collectPlaces = ({
  blocks,
  survivors,
  above,
  inCell,
  out,
}: PlacesOptions): Map<string, Place> => {
  for (const { item: block, level } of levelled({
    items: flatBlocks(blocks),
    blocksOf: (item) => [item],
    survivors,
  })) {
    if (block.type === "table") {
      for (const { item: row, level: rowLevel } of levelled({
        items: block.rows,
        blocksOf: ({ cells }) => cells.flatMap(({ content }) => content),
        survivors,
      }))
        for (const { item: cell, level: cellLevel } of levelled({
          items: row.cells,
          blocksOf: ({ content }) => content,
          survivors,
        }))
          collectPlaces({
            blocks: cell.content,
            survivors,
            above: [...above, level, rowLevel, cellLevel],
            inCell: true,
            out,
          });
      continue;
    }
    if (block.type !== "paragraph") continue;
    const place = [...above, level];
    if (inCell && block.sectionProperties !== undefined)
      out.set(block.paraId === undefined ? "" : idKey(block.paraId), place);
    for (const { item: content, level: boxLevel } of levelled({
      items: boxContents(block),
      blocksOf: (item) => item,
      survivors,
    }))
      collectPlaces({
        blocks: content,
        survivors,
        above: [...place, boxLevel],
        inCell,
        out,
      });
  }
  return out;
};

/** Whether what stands at a level has the same neighbours before it, of those both documents hold there. */
const sameLevel = (left: Level, right: Level): boolean => {
  const { anchor } = left;
  if (anchor === undefined || anchor !== right.anchor) return false;
  const before = (level: Level, other: Level): Set<string> => {
    const shared = new Set(other.siblings);
    return new Set(
      level.siblings
        .slice(0, level.siblings.indexOf(anchor))
        .filter((sibling) => shared.has(sibling)),
    );
  };
  const ahead = before(left, right);
  const now = before(right, left);
  return (
    ahead.size === now.size && [...ahead].every((sibling) => now.has(sibling))
  );
};

const samePlace = (left: Place, right: Place): boolean =>
  left.length === right.length &&
  left.every((level, index) => {
    const other = right[index];
    return other !== undefined && sameLevel(level, other);
  });

/**
 * Whether after an edit a section stands in a table cell where none stood
 * before: on another paragraph, in another cell, or past a paragraph, cell,
 * row or block it did not stand past. An edit of the paragraph's content, of
 * what else its cell holds or of the section's own properties leaves it where
 * it stands.
 */
export const gainsSectionInTableCell = (
  before: Document,
  after: Document,
): boolean =>
  documentStories(after).some((story) => {
    const next = findStoryBody(after, story)?.content ?? [];
    if (!cellsEndSection(next)) return false;
    const previous = findStoryBody(before, story)?.content ?? [];
    if (next === previous) return false;
    if (!cellsEndSection(previous)) return true;
    const ids = collectIds(previous, new Set());
    const survivors = new Set(
      [...collectIds(next, new Set())].filter((id) => ids.has(id)),
    );
    const placesIn = (blocks: readonly BlockContent[]) =>
      collectPlaces({
        blocks,
        survivors,
        above: [],
        inCell: false,
        out: new Map(),
      });
    const stood = placesIn(previous);
    return [...placesIn(next)].some(([id, place]) => {
      const was = stood.get(id);
      return was === undefined || !samePlace(was, place);
    });
  });
