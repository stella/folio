/**
 * Operations on text-box paragraphs.
 *
 * A text box's paragraphs share the package's paragraph id space, so an
 * operation names them through the package story the box is drawn in. An
 * operation whose paragraphs all sit in one text box is applied to that box's
 * own story ({@link TextBoxStory}); one that resolves or restates revisions is
 * applied to each story holding some of its ids, the innermost boxes first.
 * What an applied operation hands back (its inverse) names the package story
 * again, so it is routed the same way when it runs.
 */
import type { BlockContent, Document, Paragraph } from "../model/document";
import { type ParagraphLocation, storyParagraphs } from "./blocks";
import { IDENTITY_SPACES, identityKeysIn, idKey, slotKey } from "./ids";
import {
  findStoryParagraph,
  locateParagraphInBlockList,
} from "./paragraphLocator";
import {
  blocksEnding,
  findStoryBody,
  packageStoryOf,
  replaceStoryBody,
  sameStory,
  type StoryAddress,
  storyBody,
  textBoxStoriesIn,
  type TextBoxStory,
} from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
} from "./types";

/** Fields naming a paragraph the operation addresses, as opposed to one it mints. */
const ADDRESS_FIELDS: ReadonlySet<string> = new Set([
  "blockId",
  "beforeBlockId",
  "nextBlockId",
]);

/** Fields holding new content, whose paragraph ids are minted rather than addressed. */
const PAYLOAD_FIELDS: ReadonlySet<string> = new Set([
  "blocks",
  "table",
  "row",
  "paragraphs",
  "cellContent",
  "comment",
  "content",
]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** The paragraph ids a value (an operation or an intent) addresses, and those of the records it expects. */
const addressedParagraphIds = (value: unknown): string[] => {
  const out: string[] = [];
  const expected = (child: unknown): void => {
    if (Array.isArray(child)) {
      for (const item of child) expected(item);
      return;
    }
    if (!isRecord(child)) return;
    for (const [key, nested] of Object.entries(child)) {
      if (key === "paraId" && typeof nested === "string") out.push(nested);
      expected(nested);
    }
  };
  const visit = (child: unknown): void => {
    if (Array.isArray(child)) {
      for (const item of child) visit(item);
      return;
    }
    if (!isRecord(child)) return;
    for (const [key, nested] of Object.entries(child)) {
      if (ADDRESS_FIELDS.has(key) && typeof nested === "string")
        out.push(nested);
      else if (key === "blockIds" && Array.isArray(nested))
        for (const id of nested) {
          if (typeof id === "string") out.push(id);
        }
      else if (key === "expected") expected(nested);
      else if (!PAYLOAD_FIELDS.has(key)) visit(nested);
    }
  };
  visit(value);
  return out;
};

/**
 * A value with each `story` field `replace` gives another story for
 * restated; the parts it leaves are shared.
 */
const restated = (
  value: unknown,
  replace: (story: StoryAddress) => StoryAddress | undefined,
): unknown => {
  if (Array.isArray(value)) {
    let changed = false;
    const items = value.map((item: unknown) => {
      const next = restated(item, replace);
      changed ||= next !== item;
      return next;
    });
    return changed ? items : value;
  }
  if (!isRecord(value)) return value;
  let changed = false;
  const out: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value)) {
    // A story field holds a story address; nothing else is spelt so.
    const next =
      key === "story" && (typeof child === "string" || isRecord(child))
        ? (replace(child as StoryAddress) ?? child)
        : restated(child, replace);
    changed ||= next !== child;
    out[key] = next;
  }
  return changed ? out : value;
};

/**
 * An operation addressed to a text box's story. The operation types name a
 * package story; this one names a box's, and only applying it reads it.
 */
export const addressedTo = <Value>(
  value: Value,
  from: OpStory,
  target: StoryAddress,
): Value =>
  sameStory(target, from)
    ? value
    : (restated(value, (story) =>
        sameStory(story, from) ? target : undefined,
      ) as Value);

/** What applying or compiling hands back, naming package stories only. */
export const namingPackageStories = <Value>(value: Value): Value =>
  restated(value, (story) => {
    const named = packageStoryOf(story);
    return named === story ? undefined : named;
  }) as Value;

/** The text box among `boxes` holding the paragraph, the innermost when several do. */
const boxHolding = (
  document: Document,
  boxes: readonly TextBoxStory[],
  id: string,
): TextBoxStory | undefined =>
  boxes.findLast(
    (box) => findStoryParagraph(storyBody(document, box), id) !== undefined,
  );

const revisionKey = (id: number): string =>
  slotKey({ space: IDENTITY_SPACES.REVISION, id });

/**
 * The stories holding each revision id: the innermost text box carrying a
 * record of it, or `story` itself. Ids no box carries stay with `story`.
 */
const revisionHolders = (
  document: Document,
  story: OpStory,
  boxes: readonly TextBoxStory[],
  ids: readonly number[],
): Map<number, StoryAddress> => {
  const held = new Map<number, StoryAddress>(ids.map((id) => [id, story]));
  for (const box of boxes) {
    const keys = new Set(identityKeysIn(storyBody(document, box).content));
    for (const id of ids) if (keys.has(revisionKey(id))) held.set(id, box);
  }
  return held;
};

/**
 * Split a revision operation by the stories holding its ids: each text box
 * first, the innermost before those drawing them, then `story` itself.
 */
const splitByRevision = <Item>(options: {
  document: Document;
  story: OpStory;
  items: readonly Item[];
  idOf: (item: Item) => number;
  build: (story: StoryAddress, items: Item[]) => DocumentOp;
}): TextBoxRoute[] | undefined => {
  const { document, story, items, idOf, build } = options;
  const boxes = textBoxStoriesIn(document, story);
  if (boxes.length === 0) return undefined;
  const held = revisionHolders(
    document,
    story,
    boxes,
    items.map((item) => idOf(item)),
  );
  const own = items.filter((item) => held.get(idOf(item)) === story);
  if (own.length === items.length) return undefined;
  const routes: TextBoxRoute[] = [];
  for (const box of boxes.toReversed()) {
    const inBox = items.filter((item) => held.get(idOf(item)) === box);
    if (inBox.length > 0) routes.push(routeOf(build(box, inBox), box));
  }
  if (own.length > 0) routes.push(routeOf(build(story, own), story));
  return routes;
};

/** An operation as it is applied, and the text box story it is applied to. */
export type TextBoxRoute = { op: DocumentOp; box: TextBoxStory | undefined };

const routeOf = (op: DocumentOp, target: StoryAddress): TextBoxRoute => ({
  op,
  box:
    target === OP_STORIES.MAIN || target.kind !== "textBox"
      ? undefined
      : target,
});

/**
 * The operations an operation on `story` is applied as when some of what it
 * names sits in a text box drawn there, or `undefined` when it is applied as
 * it is. A text operation whose paragraphs sit in one box goes to that box;
 * one naming paragraphs of several stories stays, and finds no paragraph.
 */
export const textBoxRoute = (
  document: Document,
  op: DocumentOp,
  story: OpStory | undefined,
): TextBoxRoute[] | undefined => {
  if (story === undefined) return undefined;
  switch (op.type) {
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return splitByRevision({
        document,
        story,
        items: op.revisionIds,
        idOf: (id) => id,
        build: (target, revisionIds) =>
          addressedTo({ ...op, revisionIds }, story, target),
      });
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
      return splitByRevision({
        document,
        story,
        items: op.infos,
        idOf: ({ id }) => id,
        build: (target, infos) => addressedTo({ ...op, infos }, story, target),
      });
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      if (op.anchor.kind === "revision") {
        const revisionId = op.anchor.revisionId;
        return splitByRevision({
          document,
          story,
          items: [revisionId],
          idOf: (id) => id,
          build: (target) => addressedTo(op, story, target),
        });
      }
      break;
    default:
      break;
  }
  const box = textBoxHoldingAll(document, story, addressedParagraphIds(op));
  return box === undefined
    ? undefined
    : [routeOf(addressedTo(op, story, box), box)];
};

/**
 * The text box drawn in `story` that holds every one of the paragraphs, or
 * `undefined` when one of them is the story's own, or they are in several.
 */
const textBoxHoldingAll = (
  document: Document,
  story: OpStory,
  named: readonly string[],
): TextBoxStory | undefined => {
  const ids = [...new Set(named.map(idKey))];
  if (ids.length === 0) return undefined;
  const body = findStoryBody(document, story);
  if (
    body === undefined ||
    ids.some((id) => findStoryParagraph(body, id) !== undefined)
  )
    return undefined;
  const boxes = textBoxStoriesIn(document, story);
  const [first, ...rest] = ids;
  const box =
    first === undefined ? undefined : boxHolding(document, boxes, first);
  return box !== undefined &&
    rest.every((id) => {
      const holder = boxHolding(document, boxes, id);
      return holder !== undefined && sameStory(holder, box);
    })
    ? box
    : undefined;
};

/** The one story every `story` field of a value names, if they agree. */
const storyNamed = (value: unknown): StoryAddress | undefined => {
  let found: StoryAddress | undefined;
  let agreed = true;
  const visit = (child: unknown): void => {
    if (!agreed) return;
    if (Array.isArray(child)) {
      for (const item of child) visit(item);
      return;
    }
    if (!isRecord(child)) return;
    for (const [key, nested] of Object.entries(child)) {
      if (key === "story" && (typeof nested === "string" || isRecord(nested))) {
        const story = nested as StoryAddress;
        if (found === undefined) found = story;
        else if (!sameStory(found, story)) agreed = false;
      } else if (!PAYLOAD_FIELDS.has(key)) visit(nested);
    }
  };
  visit(value);
  return agreed ? found : undefined;
};

/**
 * The text box an editor intent's paragraphs all sit in, when it names them
 * through the package story drawing the box, with that story.
 */
export const textBoxOfIntent = (
  document: Document,
  intent: unknown,
): { story: OpStory; box: TextBoxStory } | undefined => {
  const named = storyNamed(intent);
  if (named === undefined) return undefined;
  const story = packageStoryOf(named);
  // An intent already naming a box's story is compiled there as it is.
  if (story !== named || findStoryBody(document, story) === undefined)
    return undefined;
  const box = textBoxHoldingAll(document, story, addressedParagraphIds(intent));
  return box === undefined ? undefined : { story, box };
};

/** A paragraph a package story names, and the story or text box holding it. */
export type LocatedThroughTextBoxes = {
  story: StoryAddress;
  location: ParagraphLocation;
};

/**
 * The paragraph `blockId` names in `story`: one of the story's own, or one
 * in a text box drawn there.
 */
export const locateThroughTextBoxes = (
  document: Document,
  story: OpStory,
  blockId: string,
): LocatedThroughTextBoxes | undefined => {
  const body = findStoryBody(document, story);
  if (body === undefined) return undefined;
  const own = findStoryParagraph(body, blockId);
  if (own !== undefined) return { story, location: own };
  for (const box of textBoxStoriesIn(document, story)) {
    const location = findStoryParagraph(storyBody(document, box), blockId);
    if (location !== undefined) return { story: box, location };
  }
  return undefined;
};

/**
 * The content of each text box drawn in `story`, at any depth, a box before
 * the boxes drawn in it.
 */
export const textBoxContentsIn = (
  document: Document,
  story: OpStory,
): readonly (readonly BlockContent[])[] =>
  findStoryBody(document, story) === undefined
    ? []
    : textBoxStoriesIn(document, story).map(
        (box) => storyBody(document, box).content,
      );

type CloseTextBoxesOptions = {
  document: Document;
  story: OpStory;
  /** A new empty paragraph, with an id of its own. */
  closing: () => Paragraph;
};

/**
 * `document` with an empty paragraph after the blocks of each text box drawn
 * in `story`, nested ones included, that end without one: a text box is a
 * container, and a container ends with a paragraph.
 */
export const withTextBoxesClosed = ({
  document,
  story,
  closing,
}: CloseTextBoxesOptions): Document => {
  let current = document;
  // A box is found again after each change: closing one rebuilds its host.
  for (let index = 0; ; index += 1) {
    const box = textBoxStoriesIn(current, story).at(index);
    if (box === undefined) return current;
    const body = storyBody(current, box);
    if (blocksEnding(body.content) === "paragraph") continue;
    current = replaceStoryBody({
      document: current,
      story: box,
      body: { ...body, content: [...body.content, closing()] },
    });
  }
};

/** A story's own paragraphs, then those of each text box drawn in it. */
export const paragraphsThroughTextBoxes = (
  document: Document,
  story: OpStory,
): ParagraphLocation[] => {
  const body = findStoryBody(document, story);
  if (body === undefined) return [];
  return [
    ...storyParagraphs(body),
    ...textBoxStoriesIn(document, story).flatMap((box) =>
      storyParagraphs(storyBody(document, box)),
    ),
  ];
};

/**
 * The paragraph `blockId` names in `story` and its physical siblings: one of
 * the story's own, or one in a text box drawn there.
 */
export const locateParagraphThroughTextBoxes = (
  document: Document,
  story: OpStory,
  blockId: string,
): ReturnType<typeof locateParagraphInBlockList> => {
  const body = findStoryBody(document, story);
  if (body === undefined) return undefined;
  const own = locateParagraphInBlockList(body, blockId);
  if (own !== undefined) return own;
  for (const box of textBoxStoriesIn(document, story)) {
    const located = locateParagraphInBlockList(
      storyBody(document, box),
      blockId,
    );
    if (located !== undefined) return located;
  }
  return undefined;
};
