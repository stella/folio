import { panic } from "better-result";

import type { FolioContentParagraphIndentation } from "../compare/content-types";
import type { ParagraphFormatting } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

type ModelIndentation = Pick<ParagraphFormatting, keyof FolioContentParagraphIndentation>;
type IndentationModelAgreement = [FolioContentParagraphIndentation, ModelIndentation] extends [
  ModelIndentation,
  FolioContentParagraphIndentation,
]
  ? unknown
  : never;

const DIRECT_PARAGRAPH_INDENTATION_KEYS = [
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
] as const satisfies readonly (keyof FolioContentParagraphIndentation)[] &
  IndentationModelAgreement;

/** The authored `w:ind` attribute cluster. */
export type DirectParagraphIndentation = FolioContentParagraphIndentation;

type CopyIndentationValueOptions = {
  target: DirectParagraphIndentation;
  key: (typeof DIRECT_PARAGRAPH_INDENTATION_KEYS)[number];
  value: number | boolean | null | undefined;
};

const copyIndentationValue = ({ target, key, value }: CopyIndentationValueOptions): void => {
  switch (key) {
    case "hangingIndent":
      if (typeof value === "boolean") {
        target.hangingIndent = value;
      }
      return;
    case "indentLeft":
    case "indentRight":
    case "indentFirstLine":
      if (typeof value === "number") {
        target[key] = value;
      }
      return;
    default: {
      const unreachable: never = key;
      return panic("Unhandled paragraph indentation property", { key: unreachable });
    }
  }
};

/** Project only authored `w:ind` attributes, preserving zero and false. */
export const paragraphIndentationFromFormatting = (
  formatting: ParagraphFormatting | null | undefined,
): DirectParagraphIndentation | undefined => {
  if (!formatting) {
    return undefined;
  }
  const indentation: DirectParagraphIndentation = {};
  for (const key of DIRECT_PARAGRAPH_INDENTATION_KEYS) {
    copyIndentationValue({ target: indentation, key, value: formatting[key] });
  }
  // `w:firstLine` and `w:hanging` are mutually exclusive alternatives. A
  // direct first-line value, including zero, clears any hanging indent a list
  // level or paragraph style supplied.
  if (indentation.indentFirstLine !== undefined && indentation.hangingIndent === undefined) {
    indentation.hangingIndent = false;
  }
  return Object.keys(indentation).length > 0 ? indentation : undefined;
};

/** Read direct indentation without materializing values inherited from a style. */
export const directParagraphIndentation = (
  attrs: ParagraphAttrs,
): DirectParagraphIndentation | undefined => {
  const original = paragraphIndentationFromFormatting(attrs._originalFormatting);
  if (!original) {
    return undefined;
  }
  const indentation: DirectParagraphIndentation = {};
  for (const key of DIRECT_PARAGRAPH_INDENTATION_KEYS) {
    if (original[key] === undefined) {
      continue;
    }
    copyIndentationValue({ target: indentation, key, value: attrs[key] });
  }
  return Object.keys(indentation).length > 0 ? indentation : undefined;
};

export const paragraphIndentationEqual = (
  left: DirectParagraphIndentation | null | undefined,
  right: DirectParagraphIndentation | null | undefined,
): boolean => DIRECT_PARAGRAPH_INDENTATION_KEYS.every((key) => left?.[key] === right?.[key]);

/** Replace the complete direct `w:ind` cluster in canonical formatting. */
export const withDirectParagraphIndentation = (
  formatting: ParagraphFormatting | null | undefined,
  indentation: DirectParagraphIndentation | null | undefined,
): ParagraphFormatting | undefined => {
  const result: ParagraphFormatting = { ...formatting };
  for (const key of DIRECT_PARAGRAPH_INDENTATION_KEYS) {
    Reflect.deleteProperty(result, key);
  }
  if (indentation) {
    Object.assign(result, indentation);
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

type ParagraphIndentationAttrPatchOptions = {
  direct: DirectParagraphIndentation | null | undefined;
  inherited: DirectParagraphIndentation | null | undefined;
};

/** Project direct-over-inherited indentation into paragraph attrs. */
export const paragraphIndentationAttrPatch = ({
  direct,
  inherited,
}: ParagraphIndentationAttrPatchOptions): Record<string, unknown> => {
  const effective = { ...inherited, ...direct };
  const hangingIndent =
    direct?.indentFirstLine === undefined
      ? (effective.hangingIndent ?? false)
      : direct.hangingIndent === true;
  return {
    indentLeft: effective.indentLeft ?? null,
    indentRight: effective.indentRight ?? null,
    indentFirstLine: effective.indentFirstLine ?? null,
    hangingIndent,
  };
};
