import { panic } from "better-result";
import {
  type ParagraphNumberingOverride,
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
} from "@stll/docx-core/model";

import { isBulletLevel, type NumberingMap } from "../docx/numberingParser";

/**
 * Which list the selection sits in, for a toolbar.
 *
 * One union, in core, for both adapters. There used to be three declarations
 * of this — React's `ui/ListButtons`, Vue's `utils/listState`, and a third in
 * Vue's editor types whose `type` had widened to `string` — none importing
 * another and none covered by a parity gate, so they drifted without anything
 * able to see it.
 *
 * The union also retires the `isInList` flag, which restated `type !== "none"`,
 * and the `numId` that was meaningful only when it was set: a state such as
 * `{ type: "none", isInList: true, numId: 7, level: -1 }` was representable and
 * meant nothing.
 */
export type ListState =
  | { readonly type: "none" }
  | { readonly type: "bullet"; readonly level: number; readonly numId?: number }
  | { readonly type: "numbered"; readonly level: number; readonly numId?: number };

/** The kinds of list a toolbar distinguishes. */
export type ListType = ListState["type"];

/** The selection is in no list. */
export const NO_LIST_STATE = { type: "none" } as const satisfies ListState;

/** A selection that is in a list, so it has a level. */
export type ActiveListState = Exclude<ListState, { type: "none" }>;

/**
 * Whether the selection is in a list at all: the one reader over the
 * discriminator, so no consumer spells the empty arm itself.
 */
export const isInListState = (state: ListState | undefined): state is ActiveListState =>
  // oxlint-disable-next-line folio-reserved-values/no-bare-reserved-compare -- `ListState.type` is this union's own discriminator, not an OOXML attribute; this is the reader every consumer calls.
  state !== undefined && state.type !== "none";

/** The list level the selection sits at, zero when it is in no list. */
export const listStateLevel = (state: ListState | undefined): number =>
  isInListState(state) ? state.level : 0;

/**
 * The toolbar's view of a paragraph's stated numbering.
 *
 * Which kind of list it is comes from the numbering definitions, through the
 * registry's own reader: a level whose `w:numFmt` is `bullet` is a bullet list
 * and anything else is a numbered one. The adapters used to decide it from the
 * id — `numId === 1` meant bullets — which was true only of documents Folio
 * had created itself, and read a hundred-item bulleted import as numbered.
 *
 * A cancellation is not a list. A level stated without an id inherits the id
 * from a tier this reader cannot see, so it keeps its level and states no id;
 * with no definition to resolve, it falls back to numbered, which is the
 * format `w:numFmt` itself defaults to.
 */
export const resolveListState = (
  numbering: NumberingMap | null | undefined,
  numPr: ParagraphNumberingOverride | undefined,
): ListState => {
  if (numPr === undefined) {
    return NO_LIST_STATE;
  }
  const level = paragraphNumberingLevel(numPr);
  if (level === undefined) {
    return NO_LIST_STATE;
  }
  const numId = paragraphNumberingReferenceId(numPr);
  if (numId === undefined) {
    return { type: "numbered", level };
  }
  const definition = numbering?.getLevel(numId, level) ?? null;
  const type = definition !== null && isBulletLevel(definition) ? "bullet" : "numbered";
  return { type, level, numId };
};

/** Whether two toolbar list states say the same thing. */
export const sameListState = (
  left: ListState | undefined,
  right: ListState | undefined,
): boolean => {
  if (left === undefined || right === undefined) {
    return left === right;
  }
  switch (left.type) {
    case "none":
      return right.type === left.type;
    case "bullet":
    case "numbered":
      return left.type === right.type && left.level === right.level && left.numId === right.numId;
    default: {
      const unhandled: never = left;
      return panic(`Unhandled list state ${JSON.stringify(unhandled)}`);
    }
  }
};
