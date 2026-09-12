import type { ParagraphFormatting } from "../types/document";
import {
  mergeParagraphFormatting,
  PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS,
} from "../utils/paragraphFormattingMerge";
import type { StyleEngine } from "./styleEngine";

/**
 * Paragraph fields currently modeled from an enclosing table style. Fields
 * not named here remain a typed refusal in the total disposition map below.
 */
export type TableParagraphPresentationOverlay = Pick<
  ParagraphFormatting,
  "spaceBefore" | "spaceAfter" | "lineSpacing" | "lineSpacingRule" | "contextualSpacing" | "frame"
>;

export type ParagraphFormattingProjectionDisposition =
  | "effective-presentation"
  | "effective-table-overlay"
  | "authored-provenance"
  | "authored-selector"
  | "run-presentation"
  | "typed-unsupported";

/**
 * One exhaustive owner for every modeled `w:pPr` field.
 *
 * `effective-table-overlay` is the deliberately supported subset of table
 * style paragraph properties. Every other presentation field still resolves
 * through document defaults, paragraph styles, and direct formatting, but a
 * table style cannot silently begin contributing it until this map changes.
 */
export const PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS = Object.freeze({
  alignment: "effective-presentation",
  bidi: "effective-presentation",
  kinsoku: "effective-presentation",
  overflowPunctuation: "effective-presentation",
  spaceBefore: "effective-table-overlay",
  spaceAfter: "effective-table-overlay",
  lineSpacing: "effective-table-overlay",
  lineSpacingRule: "effective-table-overlay",
  snapToGrid: "effective-presentation",
  beforeAutospacing: "effective-presentation",
  afterAutospacing: "effective-presentation",
  spacingExplicit: "authored-provenance",
  indentLeft: "effective-presentation",
  indentRight: "effective-presentation",
  indentFirstLine: "effective-presentation",
  hangingIndent: "effective-presentation",
  borders: "effective-presentation",
  shading: "effective-presentation",
  tabs: "effective-presentation",
  keepNext: "effective-presentation",
  keepLines: "effective-presentation",
  widowControl: "effective-presentation",
  pageBreakBefore: "effective-presentation",
  contextualSpacing: "effective-table-overlay",
  numPr: "effective-presentation",
  numPrFromStyle: "authored-provenance",
  outlineLevel: "effective-presentation",
  styleId: "authored-selector",
  frame: "effective-table-overlay",
  suppressLineNumbers: "typed-unsupported",
  suppressAutoHyphens: "effective-presentation",
  runProperties: "run-presentation",
  runInWithNext: "effective-presentation",
} as const satisfies Record<keyof ParagraphFormatting, ParagraphFormattingProjectionDisposition>);

type FieldsWithDisposition<Disposition extends ParagraphFormattingProjectionDisposition> = {
  [Field in keyof typeof PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS]: (typeof PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS)[Field] extends Disposition
    ? Field
    : never;
}[keyof typeof PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS];

type EffectiveParagraphPresentationField = FieldsWithDisposition<
  "effective-presentation" | "effective-table-overlay"
>;

const paragraphFormattingDispositionByField = new Map<
  string,
  ParagraphFormattingProjectionDisposition
>(Object.entries(PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS));

/** Paragraph properties whose resolved values affect the imported presentation. */
export type EffectiveParagraphPresentation = Readonly<
  Pick<ParagraphFormatting, EffectiveParagraphPresentationField>
>;

/** The exact modeled fields the paragraph serializer may author in `w:pPr`. */
export type AuthoredParagraphFormatting = Readonly<
  Omit<ParagraphFormatting, "numPrFromStyle" | "spacingExplicit">
>;

export type ParagraphPresentationUnsupportedProperty = {
  readonly source: "direct" | "inherited" | "table-style";
  readonly field: keyof ParagraphFormatting;
  readonly value: Exclude<ParagraphFormatting[keyof ParagraphFormatting], undefined>;
};

export type TableParagraphPresentationProjection = {
  readonly overlay?: TableParagraphPresentationOverlay;
  readonly unsupported: readonly ParagraphPresentationUnsupportedProperty[];
};

const ownsEffectivePresentation = (
  disposition: ParagraphFormattingProjectionDisposition | undefined,
): boolean => disposition === "effective-presentation" || disposition === "effective-table-overlay";

const projectFormatting = (
  formatting: ParagraphFormatting | undefined,
  includes: (disposition: ParagraphFormattingProjectionDisposition | undefined) => boolean,
): ParagraphFormatting => {
  const projected = formatting ? { ...formatting } : {};
  for (const field of Object.keys(projected)) {
    const disposition = paragraphFormattingDispositionByField.get(field);
    if (!includes(disposition) || Reflect.get(projected, field) === undefined) {
      Reflect.deleteProperty(projected, field);
    }
  }
  return projected;
};

const freezeRecursively = (value: unknown): void => {
  if (typeof value !== "object" || value === null || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freezeRecursively(child);
  Object.freeze(value);
};

const ownedFormatting = <Formatting extends ParagraphFormatting>(
  formatting: Formatting,
): Readonly<Formatting> => {
  const owned = structuredClone(formatting);
  freezeRecursively(owned);
  return owned;
};

const sameNumPr = (
  left: ParagraphFormatting["numPr"],
  right: ParagraphFormatting["numPrFromStyle"],
): boolean =>
  left !== undefined &&
  right !== undefined &&
  left.numId === right.numId &&
  (left.ilvl ?? 0) === (right.ilvl ?? 0);

/**
 * Project what the paragraph serializer can actually write from the live
 * document model. Import-only provenance never masquerades as authorship, and
 * style-sourced numbering stays inherited while it still equals its recorded
 * style value.
 */
export const projectAuthoredParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): AuthoredParagraphFormatting => {
  const projected = projectFormatting(
    formatting,
    (disposition) => disposition !== "authored-provenance",
  );
  if (sameNumPr(formatting?.numPr, formatting?.numPrFromStyle)) {
    Reflect.deleteProperty(projected, "numPr");
  }
  return ownedFormatting(projected);
};

const ownColor = <Color extends object>(color: Color | undefined): Color | undefined =>
  color === undefined ? undefined : Object.freeze({ ...color });

const ownEffectivePresentation = (
  formatting: ParagraphFormatting | undefined,
): EffectiveParagraphPresentation => {
  const owned = projectFormatting(formatting, ownsEffectivePresentation);
  if (owned.borders !== undefined) {
    const borders = { ...owned.borders };
    for (const [side, border] of Object.entries(borders)) {
      if (border === undefined) {
        Reflect.deleteProperty(borders, side);
        continue;
      }
      const color = ownColor(border.color);
      Reflect.set(
        borders,
        side,
        Object.freeze({ ...border, ...(color !== undefined && { color }) }),
      );
    }
    owned.borders = Object.freeze(borders);
  }
  if (owned.shading !== undefined) {
    const color = ownColor(owned.shading.color);
    const fill = ownColor(owned.shading.fill);
    owned.shading = Object.freeze({
      ...owned.shading,
      ...(color !== undefined && { color }),
      ...(fill !== undefined && { fill }),
    });
  }
  if (owned.tabs !== undefined) {
    const tabs = owned.tabs.map((tab) => Object.freeze({ ...tab }));
    owned.tabs = tabs;
    Object.freeze(tabs);
  }
  if (owned.numPr !== undefined) {
    owned.numPr = Object.freeze({ ...owned.numPr });
  }
  if (owned.frame !== undefined) {
    owned.frame = Object.freeze({ ...owned.frame });
  }
  return Object.freeze(owned);
};

const effectivePresentationFrom = (
  formatting: ParagraphFormatting | undefined,
): EffectiveParagraphPresentation => ownEffectivePresentation(formatting);

const unsupportedPropertiesFrom = (
  formatting: ParagraphFormatting | undefined,
  source: ParagraphPresentationUnsupportedProperty["source"],
): ParagraphPresentationUnsupportedProperty[] => {
  if (formatting === undefined) return [];
  const unsupported: ParagraphPresentationUnsupportedProperty[] = [];
  for (const { field } of Object.values(PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS)) {
    const disposition = PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS[field];
    const value = formatting[field];
    if (disposition === "typed-unsupported" && value !== undefined) {
      const evidence = structuredClone({ source, field, value });
      freezeRecursively(evidence);
      unsupported.push(evidence);
    }
  }
  return unsupported;
};

/**
 * Extract the table-style paragraph layer supported by the import model.
 * Unsupported table-style `w:pPr` fields are returned as typed evidence
 * rather than leaking into the cascade or disappearing silently.
 */
export const projectTableParagraphPresentation = (
  formatting: ParagraphFormatting | undefined,
): TableParagraphPresentationProjection | undefined => {
  if (formatting === undefined) return undefined;
  const overlay = projectFormatting(
    formatting,
    (disposition) => disposition === "effective-table-overlay",
  );
  const unsupported: ParagraphPresentationUnsupportedProperty[] = [];
  for (const { field } of Object.values(PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS)) {
    const disposition = PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS[field];
    const value = formatting[field];
    if (
      (disposition === "effective-presentation" || disposition === "typed-unsupported") &&
      value !== undefined
    ) {
      const evidence = structuredClone({
        source: "table-style" as const,
        field,
        value,
      });
      freezeRecursively(evidence);
      unsupported.push(evidence);
    }
  }
  if (Object.keys(overlay).length === 0 && unsupported.length === 0) return undefined;
  const ownedOverlay =
    Object.keys(overlay).length === 0 ? undefined : ownEffectivePresentation(overlay);
  return Object.freeze({
    ...(ownedOverlay !== undefined && { overlay: ownedOverlay }),
    unsupported: Object.freeze(unsupported),
  });
};

/** Merge table-style paragraph projections without discarding refusal evidence. */
export const mergeTableParagraphPresentations = (
  inherited: TableParagraphPresentationProjection | undefined,
  direct: TableParagraphPresentationProjection | undefined,
): TableParagraphPresentationProjection | undefined => {
  if (inherited === undefined) return direct;
  if (direct === undefined) return inherited;
  const overlay = mergeParagraphFormatting(inherited.overlay, direct.overlay);
  const unsupportedByField = new Map<
    keyof ParagraphFormatting,
    ParagraphPresentationUnsupportedProperty
  >();
  for (const unsupported of [...inherited.unsupported, ...direct.unsupported]) {
    unsupportedByField.set(unsupported.field, unsupported);
  }
  return Object.freeze({
    ...(overlay !== undefined && { overlay: ownEffectivePresentation(overlay) }),
    unsupported: Object.freeze([...unsupportedByField.values()]),
  });
};

type ResolveEffectiveParagraphPresentationOptions = {
  readonly authored: ParagraphFormatting | undefined;
  readonly styleResolver: Pick<StyleEngine, "resolveParagraphStyleInTable"> | null;
  readonly tableParagraphPresentation?: TableParagraphPresentationProjection;
};

export type ResolvedEffectiveParagraphPresentation = {
  /** Defaults, table overlay, and the selected paragraph style; no direct pPr. */
  readonly inherited: EffectiveParagraphPresentation;
  /** The complete supported presentation after direct pPr wins. */
  readonly effective: EffectiveParagraphPresentation;
  /** Presentation fields intentionally not modeled by this resolver. */
  readonly unsupported: readonly ParagraphPresentationUnsupportedProperty[];
};

/**
 * Resolve the supported paragraph presentation without consulting editor
 * state. The result is a point-in-time value: callers must resolve again from
 * the current authored model after an edit instead of caching it on a PM node.
 */
export const resolveEffectiveParagraphPresentation = ({
  authored,
  styleResolver,
  tableParagraphPresentation,
}: ResolveEffectiveParagraphPresentationOptions): ResolvedEffectiveParagraphPresentation => {
  const inheritedFormatting = styleResolver?.resolveParagraphStyleInTable(
    authored?.styleId,
    tableParagraphPresentation?.overlay,
  ).paragraphFormatting;
  const inherited = effectivePresentationFrom(inheritedFormatting);
  const direct = effectivePresentationFrom(authored);
  let effective = mergeParagraphFormatting(inherited, direct) ?? {};

  // Direct paragraph borders replace the inherited `w:pBdr` in the existing
  // import contract. Style layers themselves still merge individual sides.
  if (direct.borders !== undefined) {
    effective.borders = { ...direct.borders };
  }

  // A direct numId chooses a numbering instance and does not inherit the
  // style's ilvl. Conversely an ilvl-only direct child keeps a positive style
  // numId. A style-level numId=0 contributes no numbering attrs on import.
  if (direct.numPr?.numId !== undefined || inherited.numPr?.numId === 0) {
    if (direct.numPr === undefined) {
      Reflect.deleteProperty(effective, "numPr");
    } else {
      effective.numPr = { ...direct.numPr };
    }
  }

  // Explicitly removing style numbering also removes the style-owned marker
  // positioning indents. Direct indents remain authoritative; indentRight is
  // deliberately unaffected, matching the imported layout contract.
  const numberingRemoved =
    direct.numPr?.numId === 0 && inherited.numPr !== undefined && inherited.numPr.numId !== 0;
  if (numberingRemoved) {
    if (direct.indentLeft === undefined) Reflect.deleteProperty(effective, "indentLeft");
    if (direct.indentFirstLine === undefined) {
      Reflect.deleteProperty(effective, "indentFirstLine");
    }
    if (direct.hangingIndent === undefined) Reflect.deleteProperty(effective, "hangingIndent");
  }

  effective = ownEffectivePresentation(effective);

  const unsupported = [
    ...(tableParagraphPresentation?.unsupported ?? []),
    ...unsupportedPropertiesFrom(inheritedFormatting, "inherited"),
    ...unsupportedPropertiesFrom(authored, "direct"),
  ];

  return Object.freeze({
    inherited,
    effective,
    unsupported: Object.freeze(unsupported),
  });
};
