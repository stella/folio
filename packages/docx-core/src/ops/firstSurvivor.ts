import type {
  Document,
  DocumentBody,
  Paragraph,
  ParagraphMarkChange,
} from "../model/document";
import { storyParagraphs } from "./blocks";
import { idKey } from "./ids";
import { findStoryParagraph } from "./paragraphLocator";
import { reviewFieldsOf } from "./review";
import { findStoryBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type OpStory,
  type SetParagraphReviewOp,
  type TouchedBlocks,
} from "./types";

/**
 * The fields of a new mark that keeps the paragraph before it, the paragraph
 * after it taking over `former`, the mark the paragraph had: the tracked
 * mark at the end of a chain of such takeovers, none when it is no tracked
 * change.
 */
export const firstSurvivorFields = (
  former: ParagraphMarkChange | undefined,
): Pick<ParagraphMarkChange, "survivor" | "takenOverMark"> => {
  const takenOverMark =
    former?.survivor === "first" ? former.takenOverMark : former?.info.id;
  return takenOverMark === undefined
    ? { survivor: "first" }
    : { survivor: "first", takenOverMark };
};

/** The ids of the marks on `paragraphs`. */
const markIdsOf = (paragraphs: readonly Paragraph[]): Set<number> =>
  new Set(
    paragraphs.flatMap(({ pPrMark }) => (pPrMark ? [pPrMark.info.id] : [])),
  );

type StoryReleaseOptions = {
  before: DocumentBody;
  after: DocumentBody;
  story: OpStory;
  touched: TouchedBlocks;
};

const located = (body: DocumentBody, ids: readonly string[]): Paragraph[] =>
  ids.flatMap((id) => {
    const at = findStoryParagraph(body, id);
    return at === undefined ? [] : [at.paragraph];
  });

/**
 * What retiring the marks an edit removed takes from the keep-first rule:
 * the releases still to make, and the paragraphs an operation the edit
 * applied already released, which the edit modified as well.
 */
type TakenOverMarkReleases = {
  ops: SetParagraphReviewOp[];
  released: string[];
};

const NO_RELEASES: TakenOverMarkReleases = { ops: [], released: [] };

/** Removing a break retires the keep-first rule that depended on it. */
const storyReleases = ({
  before,
  after,
  story,
  touched,
}: StoryReleaseOptions): TakenOverMarkReleases => {
  const holders = located(before, [...touched.modified, ...touched.removed]);
  if (holders.every(({ pPrMark }) => pPrMark === undefined)) return NO_RELEASES;
  const kept = markIdsOf(
    located(after, [...touched.modified, ...touched.inserted]),
  );
  const retired = new Set<number>();
  for (const holder of holders) {
    const id = holder.pPrMark?.info.id;
    if (id === undefined || kept.has(id)) continue;
    retired.add(id);
  }
  if (retired.size === 0) return NO_RELEASES;
  const naming = new Set(
    storyParagraphs(before).flatMap(({ paragraph: { paraId, pPrMark } }) =>
      paraId !== undefined &&
      pPrMark?.takenOverMark !== undefined &&
      retired.has(pPrMark.takenOverMark)
        ? [idKey(paraId)]
        : [],
    ),
  );
  const releases: SetParagraphReviewOp[] = [];
  const released: string[] = [];
  for (const { paragraph } of storyParagraphs(after)) {
    const mark = paragraph.pPrMark;
    const named = mark?.takenOverMark;
    if (paragraph.paraId === undefined) continue;
    if (mark === undefined || named === undefined || !retired.has(named)) {
      if (naming.has(idKey(paragraph.paraId))) released.push(paragraph.paraId);
      continue;
    }
    const { survivor: _survivor, takenOverMark: _named, ...rest } = mark;
    const expected = reviewFieldsOf(paragraph);
    releases.push({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story,
      blockId: paragraph.paraId,
      expected,
      review: {
        ...expected,
        pPrMark: rest,
      },
    });
  }
  return { ops: releases, released };
};

type TakenOverMarkReleasesOptions = {
  before: Document;
  after: Document;
  touched: TouchedBlocks;
  story: OpStory | undefined;
};

/** Retire references to marks removed by an edit in its addressed story. */
export const takenOverMarkReleases = ({
  before,
  after,
  touched,
  story,
}: TakenOverMarkReleasesOptions): TakenOverMarkReleases => {
  if (
    story === undefined ||
    (touched.modified.length === 0 && touched.removed.length === 0)
  )
    return NO_RELEASES;
  const was = findStoryBody(before, story);
  const is = findStoryBody(after, story);
  return was === undefined || is === undefined || was === is
    ? NO_RELEASES
    : storyReleases({ before: was, after: is, story, touched });
};
