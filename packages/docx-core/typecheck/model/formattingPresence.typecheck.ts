/** Every admitted undefined patch value can also be held by the model. */
import type { ParagraphFormatting, TextFormatting } from "../../src/model/formatting";
import type { FormattingPatch } from "../../src/ops/types";

type PatchPresence<Formatting> = {
  [Key in keyof Formatting]-?: { [Field in Key]-?: undefined } extends Formatting ? true : false;
};

type AllTrue<Value extends Record<keyof Value, true>> = Value;

export type RunPresenceProof = AllTrue<PatchPresence<TextFormatting>>;
export type ParagraphPresenceProof = AllTrue<PatchPresence<ParagraphFormatting>>;

const RUN = { bold: undefined, italic: undefined } satisfies TextFormatting;
const PARAGRAPH = { numPr: undefined, alignment: undefined } satisfies ParagraphFormatting;
const RUN_PATCH = { bold: undefined } satisfies FormattingPatch<TextFormatting>;
const PARAGRAPH_PATCH = { numPr: undefined } satisfies FormattingPatch<ParagraphFormatting>;

export type FormattingPresenceProof = [
  typeof RUN,
  typeof PARAGRAPH,
  typeof RUN_PATCH,
  typeof PARAGRAPH_PATCH,
];
