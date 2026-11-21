/** Stable package-story addressing, shared by operations and editor projections. */
import { panic } from "better-result";

import type {
  BlockContent,
  Document,
  DocumentBody,
  HeaderFooter,
  Paragraph,
  ShapeContent,
  ShapeTextBody,
} from "../model/document";
import { idKey } from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "./leaves";
import { OP_STORIES, type OpStory } from "./types";

/**
 * A text box's `w:txbxContent`: a story drawn in a paragraph of another
 * story. Its paragraphs share the package's paragraph id space, so an
 * operation names them through the story that draws the box (see
 * `textBoxRouting.ts`); this address exists only while an operation is
 * applied to the box's own body.
 */
export type TextBoxStory = {
  readonly kind: "textBox";
  /** The story whose paragraph draws the box. */
  readonly host: StoryAddress;
  /** That paragraph's `w14:paraId`. */
  readonly blockId: string;
  /** Which of the paragraph's text boxes, in leaf order. */
  readonly index: number;
};

/** A package story, or a text box inside one. */
export type StoryAddress = OpStory | TextBoxStory;

/** The text box a story is, if it is one. */
export const textBoxStoryOf = (
  story: StoryAddress,
): TextBoxStory | undefined =>
  story === OP_STORIES.MAIN || story.kind !== "textBox" ? undefined : story;

/** The paragraphs drawing a text box, the innermost first. */
export const textBoxHostIds = (box: TextBoxStory): string[] => {
  const host = textBoxStoryOf(box.host);
  return [box.blockId, ...(host === undefined ? [] : textBoxHostIds(host))];
};

/** The package story a story is, or the one its text box is drawn in. */
export const packageStoryOf = (story: StoryAddress): OpStory => {
  if (story === OP_STORIES.MAIN || story.kind !== "textBox") return story;
  return packageStoryOf(story.host);
};

export const sameStory = (left: StoryAddress, right: StoryAddress): boolean => {
  if (left === OP_STORIES.MAIN || right === OP_STORIES.MAIN)
    return left === right;
  if (left.kind !== right.kind) return false;
  switch (left.kind) {
    case "textBox":
      return (
        right.kind === "textBox" &&
        idKey(left.blockId) === idKey(right.blockId) &&
        left.index === right.index &&
        sameStory(left.host, right.host)
      );
    case "header":
    case "footer":
      return (
        (right.kind === "header" || right.kind === "footer") &&
        left.rId === right.rId
      );
    case "footnote":
    case "endnote":
      return (
        (right.kind === "footnote" || right.kind === "endnote") &&
        left.id === right.id
      );
    default: {
      const unreachable: never = left;
      return unreachable;
    }
  }
};

export const findStoryBody = (
  document: Document,
  story: StoryAddress,
): DocumentBody | undefined => {
  if (story === OP_STORIES.MAIN) return document.package.document;
  switch (story.kind) {
    case "textBox": {
      const textBody = findTextBox(document, story)?.leaf.shape.textBody;
      return textBody === undefined ? undefined : textBoxBody(textBody);
    }
    case "header":
      return document.package.headers?.get(story.rId);
    case "footer":
      return document.package.footers?.get(story.rId);
    case "footnote":
      return document.package.footnotes?.find(({ id }) => id === story.id);
    case "endnote":
      return document.package.endnotes?.find(({ id }) => id === story.id);
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

export const storyBody = (
  document: Document,
  story: StoryAddress,
): DocumentBody =>
  findStoryBody(document, story) ??
  panic("An operation must address an existing document story.");

export const documentStories = (document: Document): OpStory[] => {
  const out: OpStory[] = [OP_STORIES.MAIN];
  for (const rId of document.package.headers?.keys() ?? [])
    out.push({ kind: "header", rId });
  for (const rId of document.package.footers?.keys() ?? [])
    out.push({ kind: "footer", rId });
  for (const { id } of document.package.footnotes ?? [])
    out.push({ kind: "footnote", id });
  for (const { id } of document.package.endnotes ?? [])
    out.push({ kind: "endnote", id });
  return out;
};

/** The block a container ends with, as far as a story's closing paragraph goes. */
type BlocksEnding = "paragraph" | "table" | "none";

/**
 * What `blocks` end with: their last paragraph or table, content controls and
 * custom XML read through. A control or wrapper that holds neither is passed
 * over like opaque markup and bookmark markers, so the block before it
 * decides.
 */
const blocksEnding = (blocks: readonly BlockContent[]): BlocksEnding => {
  for (let index = blocks.length - 1; index >= 0; index -= 1) {
    const block = blocks[index];
    if (block?.type === "paragraph") return "paragraph";
    if (block?.type === "table") return "table";
    if (block?.type === "blockSdt" || block?.type === "blockCustomXml") {
      const ending = blocksEnding(block.content);
      if (ending !== "none") return ending;
    }
  }
  return "none";
};

/** A separator or continuation note: the application's own, not a story a caret goes in. */
const isSeparatorNote = (document: Document, story: OpStory): boolean => {
  if (story === OP_STORIES.MAIN) return false;
  switch (story.kind) {
    case "header":
    case "footer":
      return false;
    case "footnote":
    case "endnote": {
      const notes =
        story.kind === "footnote"
          ? document.package.footnotes
          : document.package.endnotes;
      const note = notes?.find(({ id }) => id === story.id);
      return note?.noteType !== undefined && note.noteType !== "normal";
    }
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

/**
 * Whether `story` lacks the paragraph a story ends with: the one a caret goes
 * in and a reader shows after a final table. The body, headers, footers and
 * notes end with one (a table cell already reads the paragraph it implies);
 * a story that ends in a table or holds no paragraph gains an empty one. A
 * separator note is left as read.
 */
export const storyLacksClosingParagraph = (
  document: Document,
  story: OpStory,
): boolean => {
  const body = findStoryBody(document, story);
  return (
    body !== undefined &&
    blocksEnding(body.content) !== "paragraph" &&
    !isSeparatorNote(document, story)
  );
};

type ReplaceStoryBodyOptions = {
  document: Document;
  story: StoryAddress;
  body: DocumentBody;
};
export const replaceStoryBody = ({
  document,
  story,
  body,
}: ReplaceStoryBodyOptions): Document => {
  if (story === OP_STORIES.MAIN)
    return { ...document, package: { ...document.package, document: body } };
  const pkg = document.package;
  switch (story.kind) {
    case "textBox":
      return replaceTextBoxBody({ document, story, body });
    case "header":
    case "footer": {
      const parts = story.kind === "header" ? pkg.headers : pkg.footers;
      const previous = parts?.get(story.rId);
      if (!previous)
        return panic("Replacing a header/footer requires an existing part.");
      const next: HeaderFooter = { ...previous, content: body.content };
      const replaced = new Map(parts);
      replaced.set(story.rId, next);
      let sectionsChanged = false;
      const sections = pkg.document.sections?.map((section) => {
        const bound =
          story.kind === "header" ? section.headers : section.footers;
        const references =
          story.kind === "header"
            ? section.properties.headerReferences
            : section.properties.footerReferences;
        const variants =
          references?.filter(({ rId }) => rId === story.rId) ?? [];
        if (!bound || variants.length === 0) return section;
        sectionsChanged = true;
        const updated = new Map(bound);
        for (const { type } of variants) updated.set(type, next);
        return story.kind === "header"
          ? { ...section, headers: updated }
          : { ...section, footers: updated };
      });
      const main =
        sectionsChanged && sections
          ? { ...pkg.document, sections }
          : pkg.document;
      return story.kind === "header"
        ? {
            ...document,
            package: { ...pkg, document: main, headers: replaced },
          }
        : {
            ...document,
            package: { ...pkg, document: main, footers: replaced },
          };
    }
    case "footnote": {
      const footnotes = pkg.footnotes?.map((note) =>
        note.id === story.id ? { ...note, content: body.content } : note,
      );
      if (!footnotes)
        return panic("Replacing a footnote requires its collection.");
      return { ...document, package: { ...pkg, footnotes } };
    }
    case "endnote": {
      const endnotes = pkg.endnotes?.map((note) =>
        note.id === story.id ? { ...note, content: body.content } : note,
      );
      if (!endnotes)
        return panic("Replacing an endnote requires its collection.");
      return { ...document, package: { ...pkg, endnotes } };
    }
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

// ---------------------------------------------------------------------------
// Text boxes
// ---------------------------------------------------------------------------

/** A shape that holds a text body: a text box. */
type TextBoxLeaf = ShapeContent & {
  shape: ShapeContent["shape"] & { textBody: ShapeTextBody };
};

const isTextBoxLeaf = (node: InlineNode): node is TextBoxLeaf =>
  node.type === "shape" && node.shape.textBody !== undefined;

/** A paragraph's text boxes in leaf order, each with its path from the paragraph. */
export const paragraphTextBoxes = (
  paragraph: Paragraph,
): readonly PlacedTextBox[] => {
  // Paragraphs are never changed in place, so a paragraph's boxes are read once.
  const known = placedTextBoxes.get(paragraph);
  if (known !== undefined) return known;
  const out: PlacedTextBox[] = [];
  const path: number[] = [];
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const [index, node] of nodes.entries()) {
      path.push(index);
      if (isTextBoxLeaf(node)) out.push({ leaf: node, path: [...path] });
      else {
        const children = childNodes(node);
        if (children !== undefined) visit(children);
      }
      path.pop();
    }
  };
  visit(paragraph.content);
  placedTextBoxes.set(paragraph, out);
  return out;
};

type PlacedTextBox = { leaf: TextBoxLeaf; path: readonly number[] };

const placedTextBoxes = new WeakMap<Paragraph, readonly PlacedTextBox[]>();

/**
 * One body per text body, so reading a box's story twice gives the same
 * body, as reading any other story does.
 */
const textBoxBodies = new WeakMap<ShapeTextBody, DocumentBody>();

const textBoxBody = (textBody: ShapeTextBody): DocumentBody => {
  let body = textBoxBodies.get(textBody);
  if (body === undefined) {
    body = { content: textBody.content };
    textBoxBodies.set(textBody, body);
  }
  return body;
};

/** A paragraph of a block list, and the block holding it rebuilt around another. */
type HeldParagraph = {
  paragraph: Paragraph;
  put: (paragraph: Paragraph) => BlockContent;
};

const replacedAt = (
  blocks: readonly BlockContent[],
  index: number,
  block: BlockContent,
): BlockContent[] => {
  const out = [...blocks];
  out[index] = block;
  return out;
};

const holdingIn = (
  blocks: readonly BlockContent[],
  key: string,
  put: (blocks: BlockContent[]) => BlockContent,
): HeldParagraph | undefined => {
  for (const [index, block] of blocks.entries()) {
    const held = heldBy(block, key);
    if (held !== undefined)
      return {
        paragraph: held.paragraph,
        put: (paragraph) => put(replacedAt(blocks, index, held.put(paragraph))),
      };
  }
  return undefined;
};

const heldBy = (
  block: BlockContent,
  key: string,
): HeldParagraph | undefined => {
  switch (block.type) {
    case "paragraph":
      return block.paraId !== undefined && idKey(block.paraId) === key
        ? { paragraph: block, put: (paragraph) => paragraph }
        : undefined;
    case "table":
      for (const [rowIndex, row] of block.rows.entries())
        for (const [cellIndex, cell] of row.cells.entries()) {
          const held = holdingIn(cell.content, key, (content) => {
            const cells = [...row.cells];
            cells[cellIndex] = { ...cell, content };
            const rows = [...block.rows];
            rows[rowIndex] = { ...row, cells };
            return { ...block, rows };
          });
          if (held !== undefined) return held;
        }
      return undefined;
    case "blockSdt":
      return holdingIn(block.content, key, (content) => ({
        ...block,
        content,
      }));
    case "blockCustomXml":
      return holdingIn(block.content, key, (content) => ({
        ...block,
        content,
      }));
    default:
      return undefined;
  }
};

type FoundTextBox = {
  host: DocumentBody;
  top: number;
  held: HeldParagraph;
  leaf: TextBoxLeaf;
  path: readonly number[];
};

const findTextBox = (
  document: Document,
  story: TextBoxStory,
): FoundTextBox | undefined => {
  const host = findStoryBody(document, story.host);
  if (host === undefined) return undefined;
  const key = idKey(story.blockId);
  for (const [top, block] of host.content.entries()) {
    const held = heldBy(block, key);
    if (held === undefined) continue;
    const box = paragraphTextBoxes(held.paragraph).at(story.index);
    return box === undefined ? undefined : { host, top, held, ...box };
  }
  return undefined;
};

const withLeafAt = (
  nodes: readonly InlineNode[],
  path: readonly number[],
  leaf: InlineNode,
): InlineNode[] => {
  const [index, ...rest] = path;
  if (index === undefined) return panic("A leaf path names no node.");
  const node = nodes[index] ?? panic("A leaf path leaves its list.");
  const out = [...nodes];
  out[index] =
    rest.length === 0
      ? leaf
      : rebuildNode(
          node,
          withLeafAt(
            childNodes(node) ?? panic("A leaf path passes through a leaf."),
            rest,
            leaf,
          ),
        );
  return out;
};

type ReplaceTextBoxBodyOptions = {
  document: Document;
  story: TextBoxStory;
  body: DocumentBody;
};

/**
 * The document with a text box's body replaced. The host paragraph, and the
 * section view of the top-level block holding it, take the new box: a text
 * box holds no section boundary.
 */
const replaceTextBoxBody = ({
  document,
  story,
  body,
}: ReplaceTextBoxBodyOptions): Document => {
  const found =
    findTextBox(document, story) ??
    panic("Replacing a text box requires an existing box.");
  const leaf: TextBoxLeaf = {
    ...found.leaf,
    shape: {
      ...found.leaf.shape,
      textBody: { ...found.leaf.shape.textBody, content: body.content },
    },
  };
  const paragraph: Paragraph = {
    ...found.held.paragraph,
    content: asParagraphContent(
      withLeafAt(found.held.paragraph.content, found.path, leaf),
    ),
  };
  const previous =
    found.host.content[found.top] ?? panic("A text box host left its body.");
  const block = found.held.put(paragraph);
  const host: DocumentBody = {
    ...found.host,
    content: replacedAt(found.host.content, found.top, block),
  };
  if (found.host.sections !== undefined) {
    const sections = [...found.host.sections];
    const at = sections.findIndex(({ content }) => content.includes(previous));
    const section = sections[at];
    if (section !== undefined)
      sections[at] = {
        ...section,
        content: replacedAt(
          section.content,
          section.content.indexOf(previous),
          block,
        ),
      };
    host.sections = sections;
  }
  return replaceStoryBody({ document, story: story.host, body: host });
};

/** The text boxes drawn in a story's paragraphs, each followed by those inside it, in document order. */
export const textBoxStoriesIn = (
  document: Document,
  story: StoryAddress,
): TextBoxStory[] => {
  const out: TextBoxStory[] = [];
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks)
      switch (block.type) {
        case "paragraph": {
          const blockId = block.paraId;
          if (blockId === undefined) break;
          for (const index of paragraphTextBoxes(block).keys()) {
            const box: TextBoxStory = {
              kind: "textBox",
              host: story,
              blockId,
              index,
            };
            out.push(box, ...textBoxStoriesIn(document, box));
          }
          break;
        }
        case "table":
          for (const row of block.rows)
            for (const cell of row.cells) visit(cell.content);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content);
          break;
        default:
          break;
      }
  };
  visit(findStoryBody(document, story)?.content ?? []);
  return out;
};

/**
 * A paragraph's content with each text box's body left empty: what the
 * paragraph holds of its own story, its boxes' stories aside.
 */
export const contentOutsideTextBoxes = (
  paragraph: Paragraph,
): Paragraph["content"] => {
  let nodes: InlineNode[] = paragraph.content;
  for (const { leaf, path } of paragraphTextBoxes(paragraph))
    nodes = withLeafAt(nodes, path, {
      ...leaf,
      shape: {
        ...leaf.shape,
        textBody: { ...leaf.shape.textBody, content: [] },
      },
    });
  return nodes === paragraph.content
    ? paragraph.content
    : asParagraphContent(nodes);
};
