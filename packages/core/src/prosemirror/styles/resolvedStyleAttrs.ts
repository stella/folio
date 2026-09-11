/** Canonical paragraph style and numbering transitions. */

import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
  selectParagraphMarkProperties,
  type AuthoredParagraphProperties,
  type NumberingLevelIndentGeometry,
} from "@stll/docx-core/model";

import {
  computeListRendering,
  numberingLevelHasMarkerSlot,
  type NumberingMap,
} from "../../docx/numberingParser";
import { createNumberingLevelIndentProvenance } from "../../docx/numberingProvenance";
import type { ParagraphFormatting, TextFormatting } from "../../types/document";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { CLEARED_LIST_RENDERING_ATTRS } from "../listMarker";
import {
  transitionParagraphProperties,
  type ParagraphPropertyProjection,
} from "../paragraphPropertyMutation";
import {
  applyParagraphPropertyMutations,
  authoredParagraphPropertiesFromFormatting,
  expectParagraphPropertyState,
  type ParagraphPropertyMutation,
  type ParagraphPropertyProjectionContext,
} from "../paragraphPropertyState";
import type { ParagraphAttrs } from "../schema/nodes";
import {
  resolveEffectiveParagraphMarkFormatting,
  type RunStyleResolver,
} from "../runStyleFormatting";
import type { ResolvedParagraphStyle } from "./styleResolver";

type ResolvedStyleIdentity = {
  styleId: string | null;
  styleName?: string;
};

const effectiveNumbering = (
  authored: AuthoredParagraphProperties,
  inherited: AuthoredParagraphProperties,
): NonNullable<ParagraphFormatting["numPr"]> | null => {
  if (authored.numPr === undefined && inherited.numPr === undefined) {
    return null;
  }
  return { ...inherited.numPr, ...authored.numPr };
};

const numberingLevelGeometry = (
  formatting: ParagraphFormatting | undefined,
  hasMarkerSlot: boolean,
): NumberingLevelIndentGeometry => ({
  ...(formatting?.indentLeft !== undefined ? { indentLeft: formatting.indentLeft } : {}),
  ...(formatting?.indentRight !== undefined ? { indentRight: formatting.indentRight } : {}),
  ...(hasMarkerSlot && formatting?.indentFirstLine !== undefined
    ? { indentFirstLine: formatting.indentFirstLine }
    : {}),
  ...(hasMarkerSlot && formatting?.hangingIndent !== undefined
    ? { hangingIndent: formatting.hangingIndent }
    : {}),
});

const numberingIndentContext = (
  authored: AuthoredParagraphProperties,
  inherited: AuthoredParagraphProperties,
  numbering: NumberingMap | null | undefined,
): ParagraphFormatting["numberingLevelIndent"] | null => {
  const numPr = effectiveNumbering(authored, inherited);
  if (numPr?.numId === undefined || numPr.numId === 0) {
    return null;
  }
  const ilvl = numPr.ilvl ?? 0;
  const level = numbering?.getLevel(numPr.numId, ilvl);
  if (!level) {
    return null;
  }
  const hasMarkerSlot = numberingLevelHasMarkerSlot(level);
  const baseline = numberingLevelGeometry(level.pPr, hasMarkerSlot);
  const owned: NumberingLevelIndentGeometry = {};
  if (
    baseline.indentLeft !== undefined &&
    authored.indentLeft === undefined &&
    inherited.indentLeft === undefined
  ) {
    owned.indentLeft = baseline.indentLeft;
  }
  if (
    baseline.indentRight !== undefined &&
    authored.indentRight === undefined &&
    inherited.indentRight === undefined
  ) {
    owned.indentRight = baseline.indentRight;
  }
  const ownsFirstLine =
    authored.indentFirstLine === undefined &&
    authored.hangingIndent === undefined &&
    inherited.indentFirstLine === undefined &&
    inherited.hangingIndent === undefined;
  if (ownsFirstLine) {
    if (baseline.indentFirstLine !== undefined) {
      owned.indentFirstLine = baseline.indentFirstLine;
    }
    if (baseline.hangingIndent !== undefined) {
      owned.hangingIndent = baseline.hangingIndent;
    }
  }
  return (
    createNumberingLevelIndentProvenance({
      numId: numPr.numId,
      ilvl,
      baseline,
      owned,
    }) ?? null
  );
};

const paragraphMarkContext = (
  attrs: ParagraphAttrs,
  resolved: ResolvedParagraphStyle,
  styleId: string | undefined,
  styleResolver: RunStyleResolver | null | undefined,
  tableRunFormatting: TextFormatting | undefined,
): ParagraphPropertyProjectionContext["paragraphMark"] => {
  const current = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const authored = current.context.paragraphMark.authored;
  const inherited = selectParagraphMarkProperties(resolved.paragraphFormatting);
  const defaultTextFormatting = resolveEffectiveParagraphMarkFormatting({
    authored: {
      ...authored,
      runProperties: mergeTextFormatting(inherited.runProperties, authored.runProperties),
    },
    styleId,
    styleResolver,
    tableRunFormatting,
  });
  return {
    authored,
    effective: {
      ...(defaultTextFormatting === undefined ? {} : { defaultTextFormatting }),
      ...(authored.runInWithNext !== undefined
        ? { runInWithNext: authored.runInWithNext }
        : inherited.runInWithNext !== undefined
          ? { runInWithNext: inherited.runInWithNext }
          : {}),
    },
  };
};

const paragraphPropertyContext = (
  attrs: ParagraphAttrs,
  authoredPPr: AuthoredParagraphProperties,
  resolved: ResolvedParagraphStyle,
  numbering: NumberingMap | null | undefined,
  styleResolver: RunStyleResolver | null | undefined,
  tableRunFormatting: TextFormatting | undefined,
): ParagraphPropertyProjectionContext => {
  const inheritedPPr = authoredParagraphPropertiesFromFormatting(
    resolved.paragraphFormatting,
  );
  return {
    inheritedPPr,
    numberingLevelIndent: numberingIndentContext(authoredPPr, inheritedPPr, numbering),
    numPrFromStyle: inheritedPPr.numPr ?? null,
    paragraphMark: paragraphMarkContext(
      attrs,
      resolved,
      authoredPPr.styleId,
      styleResolver,
      tableRunFormatting,
    ),
    spacingInheritance: resolved.spacingInheritance ?? {},
  };
};

const authoredForStyleTransition = (
  attrs: ParagraphAttrs,
  identity: ResolvedStyleIdentity,
  transition:
    | { type: "preserve-direct" }
    | { type: "replace-style" }
    | { type: "restore-authored"; formatting: ParagraphFormatting | null | undefined },
): AuthoredParagraphProperties => {
  const current = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const authored =
    transition.type === "restore-authored"
      ? authoredParagraphPropertiesFromFormatting(transition.formatting)
      : structuredClone(current.authoredPPr);
  if (transition.type === "replace-style") {
    for (const key of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
      if (PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[key].styleTransition === "replace") {
        Reflect.deleteProperty(authored, key);
      }
    }
  }
  if (identity.styleId === null) {
    Reflect.deleteProperty(authored, "styleId");
  } else {
    authored.styleId = identity.styleId;
  }
  return authored;
};

const listRenderingAttrs = (
  authored: AuthoredParagraphProperties,
  inherited: AuthoredParagraphProperties,
  numbering: NumberingMap | null | undefined,
): Record<string, unknown> => {
  const numPr = effectiveNumbering(authored, inherited);
  return numPr?.numId === undefined || numPr.numId === 0
    ? CLEARED_LIST_RENDERING_ATTRS
    : listAttrsFromNumbering({ numId: numPr.numId, ilvl: numPr.ilvl ?? 0 }, numbering);
};

type ParagraphPropertiesForStyleTransitionOptions = {
  attrs: ParagraphAttrs;
  identity: ResolvedStyleIdentity;
  numbering: NumberingMap | null | undefined;
  resolved: ResolvedParagraphStyle;
  styleResolver?: RunStyleResolver | null;
  tableRunFormatting?: TextFormatting;
  transition:
    | { type: "preserve-direct" }
    | { type: "replace-style" }
    | { type: "restore-authored"; formatting: ParagraphFormatting | null | undefined };
};

/** Project one style transition and all dependent paragraph caches atomically. */
export const paragraphPropertiesForStyleTransition = ({
  attrs,
  identity,
  numbering,
  resolved,
  styleResolver,
  tableRunFormatting,
  transition,
}: ParagraphPropertiesForStyleTransitionOptions): ParagraphPropertyProjection => {
  const authoredPPr = authoredForStyleTransition(attrs, identity, transition);
  const context = paragraphPropertyContext(
    attrs,
    authoredPPr,
    resolved,
    numbering,
    styleResolver,
    tableRunFormatting,
  );
  const nonPropertyAttrs = {
    ...attrs,
    ...listRenderingAttrs(authoredPPr, context.inheritedPPr, numbering),
    _tableOfContentsLevel:
      identity.styleId === null
        ? null
        : tableOfContentsStyleLevel({
            styleId: identity.styleId,
            ...(identity.styleName ? { styleName: identity.styleName } : {}),
          }) ?? null,
  };
  return transitionParagraphProperties({
    attrs: nonPropertyAttrs,
    state: {
      type: "update",
      authored: { type: "replace", authoredPPr },
      context: { type: "replace", context },
    },
  });
};

/** Direct numbering value that keeps inherited numbering disabled after reload. */
export const numPrAfterListRemoval = (attrs: ParagraphAttrs): ParagraphAttrs["numPr"] | null => {
  const state = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const styleNumPr = state.context.numPrFromStyle;
  return styleNumPr?.numId !== undefined && styleNumPr.numId !== 0
    ? { numId: 0, ilvl: attrs.numPr?.ilvl ?? styleNumPr.ilvl ?? 0 }
    : null;
};

/** Remove direct numbering and its level geometry through the canonical state owner. */
export const paragraphPropertiesForListRemoval = (
  attrs: ParagraphAttrs,
): ParagraphPropertyProjection => {
  const numPr = numPrAfterListRemoval(attrs);
  const mutation: ParagraphPropertyMutation =
    numPr === null
      ? { key: "numPr", mutation: { type: "remove" } }
      : { key: "numPr", mutation: { type: "set", value: numPr } };
  const state = expectParagraphPropertyState(attrs._paragraphPropertyState);
  return transitionParagraphProperties({
    attrs: { ...attrs, ...CLEARED_LIST_RENDERING_ATTRS },
    state: {
      type: "update",
      authored: { type: "mutate", mutations: [mutation] },
      context: {
        type: "replace",
        context: { ...state.context, numberingLevelIndent: null },
      },
    },
  });
};

type ParagraphPropertiesForListLevelTransitionOptions = {
  attrs: ParagraphAttrs;
  numPr: NonNullable<ParagraphFormatting["numPr"]>;
  numbering: NumberingMap | null | undefined;
};

/** Apply direct list identity and recompute level provenance atomically. */
export const paragraphPropertiesForListLevelTransition = ({
  attrs,
  numPr,
  numbering,
}: ParagraphPropertiesForListLevelTransitionOptions): ParagraphPropertyProjection => {
  const state = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const authoredPPr = applyParagraphPropertyMutations(state.authoredPPr, [
    { key: "numPr", mutation: { type: "set", value: numPr } },
  ]);
  const context = {
    ...state.context,
    numberingLevelIndent: numberingIndentContext(
      authoredPPr,
      state.context.inheritedPPr,
      numbering,
    ),
  };
  return transitionParagraphProperties({
    attrs: {
      ...attrs,
      ...(numPr.numId === undefined
        ? CLEARED_LIST_RENDERING_ATTRS
        : listAttrsFromNumbering(
            { numId: numPr.numId, ilvl: numPr.ilvl ?? 0 },
            numbering,
          )),
      listImplicitChildLevelAdvances: attrs.listImplicitChildLevelAdvances ?? null,
    },
    state: {
      type: "update",
      authored: { type: "replace", authoredPPr },
      context: { type: "replace", context },
    },
  });
};

/** Project one resolved numbering level into non-property list-rendering attrs. */
export function listAttrsFromNumbering(
  numPr: { numId: number; ilvl: number },
  numbering: NumberingMap | null | undefined,
): Record<string, unknown> {
  const targetNumPr = { numId: numPr.numId, ilvl: numPr.ilvl };
  const rendering = numbering ? computeListRendering(targetNumPr, numbering) : null;
  return {
    ...CLEARED_LIST_RENDERING_ATTRS,
    listNumFmt: rendering?.numFmt ?? null,
    listIsBullet: rendering?.isBullet ?? null,
    listIsLegal: rendering?.isLegal ?? null,
    listMarker: rendering?.marker ?? null,
    listMarkerTemplate: rendering?.markerTemplate ?? null,
    listMarkerHidden: rendering?.markerHidden ?? null,
    listMarkerFormatting: rendering?.markerFormatting ?? null,
    listMarkerAlignment: rendering?.markerAlignment ?? null,
    listMarkerSuffix: rendering?.markerSuffix ?? null,
    listMarkerAllCaps: rendering?.markerAllCaps ?? null,
    listLevelNumFmts: rendering?.levelNumFmts ?? null,
    listLevelStarts: rendering?.levelStarts ?? null,
    listAbstractNumId: rendering?.abstractNumId ?? null,
    listStartOverride: rendering?.startOverride ?? null,
  };
}
