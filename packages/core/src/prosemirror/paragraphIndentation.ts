import {
  type NonEmptyNumberingLevelIndentGeometry,
  type NumberingLevelIndentGeometry,
} from "@stll/docx-core/model";
import { createNumberingLevelIndentProvenance } from "../docx/numberingProvenance";
import {
  transitionParagraphProperties,
  type ParagraphPropertyProjection,
} from "./paragraphPropertyMutation";
import {
  expectParagraphPropertyState,
  type ParagraphPropertyMutation,
} from "./paragraphPropertyState";
import type { ParagraphAttrs } from "./schema/nodes";

const PARAGRAPH_INDENT_SIDES = ["left", "right", "firstLine"] as const;
export type ParagraphIndentSide = (typeof PARAGRAPH_INDENT_SIDES)[number];

type DirectParagraphIndentation = {
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  hangingIndent?: boolean;
};

export type ParagraphIndentationTransition =
  | { type: "set-direct"; values: DirectParagraphIndentation }
  | { type: "force-visible-zero"; sides: readonly ParagraphIndentSide[] }
  | {
      type: "reset-to-inherited";
      inherited: DirectParagraphIndentation & {
        numberingLevelIndent?: NonEmptyNumberingLevelIndentGeometry;
      };
      sides: readonly ParagraphIndentSide[];
    };

export type ParagraphIndentationTransitionResult =
  | { type: "applied"; projection: ParagraphPropertyProjection }
  | {
      type: "unsupported";
      reason: "numbering-first-line-zero-is-not-paragraph-representable";
    };

const keysForSide = (
  side: ParagraphIndentSide,
): readonly ("indentLeft" | "indentRight" | "indentFirstLine" | "hangingIndent")[] => {
  switch (side) {
    case "left":
      return ["indentLeft"];
    case "right":
      return ["indentRight"];
    case "firstLine":
      return ["indentFirstLine", "hangingIndent"];
  }
};

const directSides = (values: DirectParagraphIndentation): ParagraphIndentSide[] => {
  const sides: ParagraphIndentSide[] = [];
  if (values.indentLeft !== undefined) {
    sides.push("left");
  }
  if (values.indentRight !== undefined) {
    sides.push("right");
  }
  if (values.indentFirstLine !== undefined || values.hangingIndent !== undefined) {
    sides.push("firstLine");
  }
  return sides;
};

const valuesForTransition = (
  transition: ParagraphIndentationTransition,
): { sides: readonly ParagraphIndentSide[]; values: DirectParagraphIndentation } => {
  switch (transition.type) {
    case "set-direct":
      return { sides: directSides(transition.values), values: transition.values };
    case "force-visible-zero": {
      const values: DirectParagraphIndentation = {};
      for (const side of transition.sides) {
        if (side === "left") {
          values.indentLeft = 0;
        } else if (side === "right") {
          values.indentRight = 0;
        } else {
          values.indentFirstLine = 0;
          values.hangingIndent = false;
        }
      }
      return { sides: transition.sides, values };
    }
    case "reset-to-inherited":
      return { sides: transition.sides, values: transition.inherited };
  }
};

const activeNumberingHasFirstLineBaseline = (attrs: ParagraphAttrs): boolean => {
  const active = attrs.numPr;
  const inherited = expectParagraphPropertyState(
    attrs._paragraphPropertyState,
  ).context.numberingLevelIndent;
  return (
    active?.numId !== undefined &&
    active.numId !== 0 &&
    inherited !== undefined &&
    inherited.numId === active.numId &&
    (inherited.ilvl ?? 0) === (active.ilvl ?? 0) &&
    inherited.baseline.indentFirstLine !== undefined &&
    inherited.baseline.indentFirstLine !== 0
  );
};

const forcesUnrepresentableFirstLineZero = (
  attrs: ParagraphAttrs,
  transition: ParagraphIndentationTransition,
): boolean => {
  if (transition.type === "reset-to-inherited" || !activeNumberingHasFirstLineBaseline(attrs)) {
    return false;
  }
  if (transition.type === "force-visible-zero") {
    return transition.sides.includes("firstLine");
  }
  return transition.values.indentFirstLine === 0 && transition.values.hangingIndent !== true;
};

/** Apply one explicit direct-indent transition and update its ownership atomically. */
export const attrsWithParagraphIndentationTransition = (
  attrs: ParagraphAttrs,
  transition: ParagraphIndentationTransition,
): ParagraphIndentationTransitionResult => {
  if (forcesUnrepresentableFirstLineZero(attrs, transition)) {
    return {
      type: "unsupported",
      reason: "numbering-first-line-zero-is-not-paragraph-representable",
    };
  }
  const { sides, values } = valuesForTransition(transition);
  const state = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const numberingLevelIndent = state.context.numberingLevelIndent;
  const ownedNumberingIndent: NumberingLevelIndentGeometry =
    numberingLevelIndent?.type === "owned" ? { ...numberingLevelIndent.owned } : {};
  const mutations: ParagraphPropertyMutation[] = [];

  for (const side of sides) {
    for (const key of keysForSide(side)) {
      mutations.push({ key, mutation: { type: "remove" } });
    }
    if (side === "left") {
      Reflect.deleteProperty(ownedNumberingIndent, "indentLeft");
    } else if (side === "right") {
      Reflect.deleteProperty(ownedNumberingIndent, "indentRight");
    } else if (side === "firstLine") {
      Reflect.deleteProperty(ownedNumberingIndent, "indentFirstLine");
      Reflect.deleteProperty(ownedNumberingIndent, "hangingIndent");
    }
  }

  if (transition.type === "reset-to-inherited") {
    const inheritedNumbering = transition.inherited.numberingLevelIndent;
    if (inheritedNumbering) {
      Object.assign(ownedNumberingIndent, inheritedNumbering);
    }
  } else {
    for (const side of sides) {
      for (const key of keysForSide(side)) {
        const value = values[key];
        if (value !== undefined) {
          mutations.push({ key, mutation: { type: "set", value } });
        }
      }
    }
  }

  const nextNumberingLevelIndent = numberingLevelIndent
    ? createNumberingLevelIndentProvenance({
        numId: numberingLevelIndent.numId,
        ilvl: numberingLevelIndent.ilvl,
        baseline: numberingLevelIndent.baseline,
        owned: ownedNumberingIndent,
      })
    : null;

  return {
    type: "applied",
    projection: transitionParagraphProperties({
      attrs,
      state: {
        type: "update",
        authored: { type: "mutate", mutations },
        context: {
          type: "replace",
          context: {
            ...state.context,
            numberingLevelIndent: nextNumberingLevelIndent ?? null,
          },
        },
      },
    }),
  };
};
