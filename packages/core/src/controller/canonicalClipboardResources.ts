import { panic, Result } from "better-result";
import {
  DOCUMENT_OP_TYPES,
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
} from "@stll/docx-core/ops";
import {
  BODY_TEXT_OUTLINE_LEVEL,
  NO_PARAGRAPH_NUMBERING,
  isThemeColor,
  themeColorSlot,
} from "@stll/docx-core/model";
import { isValidHexColor, resolveColorToHex } from "../utils/colorResolver";
import { resolveThemeFont } from "../utils/fontResolver";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";
import { createStyleEngine } from "../style-engine/styleEngine";
import { resolveStyleInheritance } from "../docx/styleParser";
import { visitParagraphRuns } from "../docx/paragraphTraversal";
import { FONT_THEME_VALUES } from "../types/documentEnumValues";
import { resolveParagraphBodyRunFormatting } from "../prosemirror/runStyleFormatting";
import { cascadeStyleTextFormatting } from "../prosemirror/styles/styleToggleCascade";
import type {
  Document,
  Paragraph,
  ParagraphFormatting,
  Style,
  StyleDefinitions,
  TextFormatting,
} from "../types/document";

const STYLE_REFERENCE_FIELDS = new Set([
  "styleId",
  "basedOn",
  "next",
  "link",
  "pStyle",
  "numStyleLink",
  "styleLink",
]);

type ReferenceVisitor = (key: string, value: unknown, owner: object) => void;

const visitReferences = (value: unknown, visitor: ReferenceVisitor): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) visitReferences(item, visitor);
    return;
  }
  for (const [key, item] of Object.entries(value)) {
    if (key === "listRendering") continue;
    visitor(key, item, value);
    visitReferences(item, visitor);
  }
};

/** Normalized HTML states effective CSS; foreign style names carry no package ownership. */
export const flattenClipboardStyleReferences = (paragraphs: readonly Paragraph[]): void => {
  visitReferences(paragraphs, (key, _value, owner) => {
    if (STYLE_REFERENCE_FIELDS.has(key)) Reflect.deleteProperty(owner, key);
  });
};

const refused = (message: string) =>
  Result.err(
    new DocumentOpRefusal({
      opType: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message,
    }),
  );

type FontFamily = NonNullable<TextFormatting["fontFamily"]>;
type ThemeFontSlot = Extract<keyof FontFamily, `${string}Theme`>;
type ConcreteFontSlot = Exclude<keyof FontFamily, ThemeFontSlot | "hint">;

const THEME_FONT_SLOTS = new Map(
  Object.entries({
    asciiTheme: "ascii",
    hAnsiTheme: "hAnsi",
    eastAsiaTheme: "eastAsia",
    csTheme: "cs",
  } as const satisfies Record<ThemeFontSlot, ConcreteFontSlot>),
);

const THEME_FONT_REFERENCES: ReadonlySet<string> = new Set(FONT_THEME_VALUES);

type LowerSourceThemeOptions = { value: unknown; source: Document };

const lowerSourceTheme = ({
  value,
  source,
}: LowerSourceThemeOptions): Result<void, DocumentOpRefusal> => {
  let unavailable: string | undefined;
  visitReferences(value, (key, item, owner) => {
    if (unavailable !== undefined) return;
    const fontSlot = THEME_FONT_SLOTS.get(key);
    if (fontSlot !== undefined && typeof item === "string") {
      const explicit = Reflect.get(owner, fontSlot);
      const resolved = THEME_FONT_REFERENCES.has(item)
        ? resolveThemeFont(item, source.package.theme?.fontScheme)
        : null;
      const font = resolved || (typeof explicit === "string" ? explicit : undefined);
      if (font === undefined || font.length === 0) {
        unavailable = `The source package does not resolve theme font ${item}.`;
        return;
      }
      Reflect.set(owner, fontSlot, font);
      Reflect.deleteProperty(owner, key);
      return;
    }
    if (key !== "themeColor") return;
    if (item === "none") return;
    if (Reflect.get(owner, "auto") === true) {
      Reflect.deleteProperty(owner, "themeColor");
      Reflect.deleteProperty(owner, "themeTint");
      Reflect.deleteProperty(owner, "themeShade");
      return;
    }
    const rgb = Reflect.get(owner, "rgb");
    const tint = Reflect.get(owner, "themeTint");
    const shade = Reflect.get(owner, "themeShade");
    if (typeof item !== "string" || !isThemeColor(item)) {
      if (typeof rgb !== "string" || !isValidHexColor(rgb)) {
        unavailable = "A source theme colour has no known slot or explicit RGB value.";
        return;
      }
      Reflect.deleteProperty(owner, "themeColor");
      Reflect.deleteProperty(owner, "themeTint");
      Reflect.deleteProperty(owner, "themeShade");
      return;
    }
    const slot = themeColorSlot(item);
    const schemeColour = slot === undefined ? undefined : source.package.theme?.colorScheme?.[slot];
    if (schemeColour === undefined || !isValidHexColor(schemeColour)) {
      unavailable = `The source package does not resolve theme colour ${item}.`;
      return;
    }
    const color = {
      themeColor: item,
      ...(typeof rgb === "string" ? { rgb } : {}),
      ...(typeof tint === "string" ? { themeTint: tint } : {}),
      ...(typeof shade === "string" ? { themeShade: shade } : {}),
    };
    const resolved = resolveColorToHex(color, source.package.theme);
    if (resolved === undefined) {
      unavailable = `The source package does not resolve theme colour ${item}.`;
      return;
    }
    Reflect.set(owner, "rgb", resolved);
    Reflect.deleteProperty(owner, "themeColor");
    Reflect.deleteProperty(owner, "themeTint");
    Reflect.deleteProperty(owner, "themeShade");
  });
  return unavailable === undefined ? Result.ok(undefined) : refused(unavailable);
};

type ImportClipboardStylesOptions = {
  destination: Document;
  source: Document;
  paragraphs: readonly Paragraph[];
};

type ImportedClipboardStyles = {
  styles: StyleDefinitions | undefined;
  paragraphs: Paragraph[];
  styleIds: ReadonlyMap<string, string>;
  numberingIds: ReadonlySet<number>;
};

type MaterializeClipboardDefaultsOptions = {
  source: Document;
  destination: Document;
  paragraphs: Paragraph[];
};

// ECMA default spacing, also used by the existing compare style importer.
const NEUTRAL_PARAGRAPH_PROPERTIES = {
  bidi: false,
  keepNext: false,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
  contextualSpacing: false,
  beforeAutospacing: false,
  afterAutospacing: false,
  hangingIndent: false,
  suppressLineNumbers: false,
  suppressAutoHyphens: false,
  runInWithNext: false,
  alignment: "start",
  indentLeft: 0,
  indentRight: 0,
  indentFirstLine: 0,
  outlineLevel: BODY_TEXT_OUTLINE_LEVEL,
  numPr: NO_PARAGRAPH_NUMBERING,
  tabs: [],
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: 240,
  lineSpacingRule: "auto",
} satisfies Partial<ParagraphFormatting>;

const NEUTRAL_RUN_PROPERTIES = {
  color: { auto: true },
  highlight: "none",
  underline: { style: "none" },
  spacing: 0,
  position: 0,
  kerning: 0,
  vertAlign: "baseline",
} satisfies Partial<TextFormatting>;

const materializeClipboardDefaults = ({
  source,
  destination,
  paragraphs,
}: MaterializeClipboardDefaultsOptions): Result<void, DocumentOpRefusal> => {
  const sourceDefinitions = structuredClone(source.package.styles);
  if (sourceDefinitions !== undefined) {
    const sourceMap = new Map(sourceDefinitions.styles.map((style) => [style.styleId, style]));
    sourceDefinitions.styles = sourceDefinitions.styles.map((style) =>
      resolveStyleInheritance(style, sourceMap),
    );
  }
  const engine = createStyleEngine(sourceDefinitions);
  const destinationDefaults = createStyleEngine(destination.package.styles).resolveParagraphStyle(
    undefined,
  );
  const neutralizeBooleanDefaults = (formatting: object, defaults: object | undefined): void => {
    for (const [key, value] of Object.entries(defaults ?? {})) {
      if (typeof value === "boolean" && Reflect.get(formatting, key) === undefined)
        Reflect.set(formatting, key, false);
    }
  };
  let unavailable: string | undefined;
  const recordRefusal = (message: string): void => {
    unavailable ??= message;
  };
  for (const paragraph of paragraphs) {
    const authored = paragraph.formatting;
    const implicitStyle = engine.getDefaultParagraphStyle();
    const styleId = authored?.styleId ?? implicitStyle?.styleId;
    const resolved = engine.resolveParagraphStyle(styleId);
    const formatting =
      mergeParagraphFormatting(resolved.paragraphFormatting, authored) ??
      (destinationDefaults.paragraphFormatting === undefined ? undefined : {});
    if (formatting !== undefined) {
      // Style numbering must remain inherited: writing it directly changes
      // OOXML precedence between level and style indentation.
      if (authored?.numPr === undefined && sourceDefinitions?.docDefaults?.pPr?.numPr === undefined)
        delete formatting.numPr;
      if (authored?.numPrFromStyle === undefined) delete formatting.numPrFromStyle;
      if (authored?.preserved !== undefined) formatting.preserved = authored.preserved;
      for (const [key, value] of Object.entries(destinationDefaults.paragraphFormatting ?? {})) {
        if (
          key === "preserved" ||
          key === "styleId" ||
          key === "runProperties" ||
          // Authored spacing provenance is not an inherited visual property.
          key === "spacingExplicit" ||
          value === undefined ||
          Reflect.get(formatting, key) !== undefined ||
          Reflect.get(resolved.paragraphFormatting ?? {}, key) !== undefined
        )
          continue;
        const neutral = Reflect.get(NEUTRAL_PARAGRAPH_PROPERTIES, key);
        if (neutral === undefined) {
          recordRefusal(`The source package does not resolve inherited paragraph property ${key}.`);
          continue;
        }
        Reflect.set(formatting, key, structuredClone(neutral));
      }
      if (
        authored?.styleId === undefined &&
        implicitStyle !== undefined &&
        engine.hasStyle(implicitStyle.styleId)
      )
        formatting.styleId = implicitStyle.styleId;
      paragraph.formatting = formatting;
    }
    const inherited = resolveParagraphBodyRunFormatting({ styleId, styleResolver: engine });
    const materializeRun = (direct: TextFormatting | undefined): TextFormatting | undefined => {
      const named = direct?.styleId;
      const cascade =
        cascadeStyleTextFormatting([
          {
            cascade:
              named === undefined ? inherited.defaultToggleCascade : inherited.baseToggleCascade,
            type: "carried",
          },
          {
            formatting: named === undefined ? undefined : engine.getRunStyleOwnProperties(named),
            type: "style",
          },
          { formatting: direct, type: "direct" },
        ]).formatting ?? {};
      neutralizeBooleanDefaults(cascade, destinationDefaults.runFormatting);
      const themeLowered = lowerSourceTheme({ value: cascade, source });
      if (themeLowered.isErr()) recordRefusal(themeLowered.error.message);
      for (const [slot, value] of Object.entries(
        destinationDefaults.runFormatting?.fontFamily ?? {},
      )) {
        const concreteSlot = THEME_FONT_SLOTS.get(slot) ?? slot;
        if (
          value === undefined ||
          Reflect.get(cascade.fontFamily ?? {}, concreteSlot) !== undefined
        )
          continue;
        if (slot === "hint") {
          cascade.fontFamily = { ...cascade.fontFamily, hint: "default" };
          continue;
        }
        // A missing source script font is application/environment dependent;
        // adopting the destination's explicit font would change the paste.
        recordRefusal(`The source package does not resolve inherited font slot ${slot}.`);
      }
      for (const [key, value] of Object.entries(destinationDefaults.runFormatting ?? {})) {
        if (
          key === "preserved" ||
          key === "styleId" ||
          value === undefined ||
          Reflect.get(cascade, key) !== undefined
        )
          continue;
        const neutral = Reflect.get(NEUTRAL_RUN_PROPERTIES, key);
        if (neutral === undefined) {
          recordRefusal(`The source package does not resolve inherited run property ${key}.`);
          continue;
        }
        Reflect.set(cascade, key, structuredClone(neutral));
      }
      if (direct?.preserved !== undefined) cascade.preserved = direct.preserved;
      return cascade;
    };
    visitParagraphRuns(paragraph, (run) => {
      const runFormatting = materializeRun(run.formatting);
      if (runFormatting !== undefined) run.formatting = runFormatting;
      for (const change of run.propertyChanges ?? []) {
        const previous = materializeRun(change.previousFormatting);
        const current = materializeRun(change.currentFormatting);
        if (previous !== undefined) change.previousFormatting = previous;
        if (current !== undefined && Object.hasOwn(change, "currentFormatting"))
          change.currentFormatting = current;
      }
    });
    for (const change of paragraph.propertyChanges ?? []) {
      const previous = change.previousFormatting;
      const before = mergeParagraphFormatting(
        engine.resolveParagraphStyle(previous?.styleId).paragraphFormatting,
        previous,
      );
      if (before !== undefined) {
        if (previous?.numPr === undefined) delete before.numPr;
        if (previous?.preserved !== undefined) before.preserved = previous.preserved;
        change.previousFormatting = before;
      }
    }
    if (paragraph.content.length === 0) {
      const markFormatting = materializeRun(paragraph.formatting?.runProperties);
      if (markFormatting !== undefined)
        paragraph.formatting = { ...paragraph.formatting, runProperties: markFormatting };
    }
  }
  return unavailable === undefined ? Result.ok(undefined) : refused(unavailable);
};

/** Import the transitive style/numbering dependencies of the clipboard, retaining destination defaults. */
export const importClipboardStyles = ({
  destination,
  source,
  paragraphs,
}: ImportClipboardStylesOptions): Result<ImportedClipboardStyles, DocumentOpRefusal> => {
  const clonedParagraphs = structuredClone([...paragraphs]);
  const materialized = materializeClipboardDefaults({
    source,
    destination,
    paragraphs: clonedParagraphs,
  });
  if (materialized.isErr()) return materialized;
  const sourceStyles = new Map<string, Style>();
  for (const style of source.package.styles?.styles ?? []) {
    if (sourceStyles.has(style.styleId))
      return refused("The source package contains duplicate style identities.");
    sourceStyles.set(style.styleId, style);
  }
  const styleQueue: string[] = [];
  const numberingQueue: number[] = [];
  const selectedStyles = new Set<string>();
  const numberingIds = new Set<number>();
  const selectReferences: ReferenceVisitor = (key, value) => {
    if (
      STYLE_REFERENCE_FIELDS.has(key) &&
      typeof value === "string" &&
      !selectedStyles.has(value)
    ) {
      selectedStyles.add(value);
      styleQueue.push(value);
    }
    if (key === "numId" && typeof value === "number" && !numberingIds.has(value)) {
      numberingIds.add(value);
      numberingQueue.push(value);
    }
  };
  for (const paragraph of clonedParagraphs) visitReferences(paragraph, selectReferences);
  const sourceNums = source.package.numbering?.nums ?? [];
  const sourceAbstracts = source.package.numbering?.abstractNums ?? [];
  const nums = new Map(sourceNums.map((num) => [num.numId, num]));
  const abstracts = new Map(sourceAbstracts.map((abstract) => [abstract.abstractNumId, abstract]));
  if (nums.size !== sourceNums.length || abstracts.size !== sourceAbstracts.length)
    return refused("The source package contains duplicate numbering identities.");
  let styleIndex = 0;
  let numberingIndex = 0;
  while (styleIndex < styleQueue.length || numberingIndex < numberingQueue.length) {
    while (styleIndex < styleQueue.length) {
      const id = styleQueue.at(styleIndex++);
      if (id === undefined) return refused("A clipboard style dependency was lost.");
      const style = sourceStyles.get(id);
      if (style === undefined)
        return refused(`The source package does not define clipboard style ${id}.`);
      visitReferences(style, selectReferences);
    }
    while (numberingIndex < numberingQueue.length) {
      const id = numberingQueue.at(numberingIndex++);
      if (id === undefined) return refused("A clipboard numbering dependency was lost.");
      const num = nums.get(id);
      if (num === undefined)
        return refused(`The source package does not define clipboard numbering ${id}.`);
      const abstract = abstracts.get(num.abstractNumId);
      if (abstract === undefined)
        return refused("A clipboard numbering instance has no abstract definition.");
      visitReferences(num, selectReferences);
      visitReferences(abstract, selectReferences);
    }
  }
  for (const id of selectedStyles) {
    const ancestors = new Set<string>();
    let ancestor: string | undefined = id;
    while (ancestor !== undefined) {
      if (ancestors.has(ancestor))
        return refused("The source clipboard style inheritance contains a cycle.");
      ancestors.add(ancestor);
      ancestor = sourceStyles.get(ancestor)?.basedOn;
    }
  }
  const occupied = new Set(destination.package.styles?.styles.map(({ styleId }) => styleId));
  // Reserve every source name before minting, so an allocated suffix cannot steal a later source name.
  const reserved = new Set([...occupied, ...selectedStyles]);
  const styleIds = new Map<string, string>();
  for (const id of [...selectedStyles].toSorted()) {
    if (!occupied.has(id)) {
      styleIds.set(id, id);
      continue;
    }
    let ordinal = 1;
    let importedId = `${id}_clipboard${ordinal}`;
    while (reserved.has(importedId)) importedId = `${id}_clipboard${++ordinal}`;
    reserved.add(importedId);
    styleIds.set(id, importedId);
  }
  const remap: ReferenceVisitor = (key, value, owner) => {
    if (!STYLE_REFERENCE_FIELDS.has(key) || typeof value !== "string") return;
    const mapped = styleIds.get(value);
    if (mapped === undefined)
      panic("A selected clipboard style has no allocated destination identity.");
    Reflect.set(owner, key, mapped);
  };
  const importedStyles: Style[] = [];
  for (const id of [...selectedStyles].toSorted()) {
    const style = sourceStyles.get(id);
    if (style === undefined) return refused("A selected clipboard style has no source definition.");
    const imported = structuredClone(style);
    visitReferences(imported, remap);
    const lowered = lowerSourceTheme({ value: imported, source });
    if (lowered.isErr()) return lowered;
    // A source default applies to its source story; importing it must not replace destination defaults.
    if (imported.default === true) delete imported.default;
    importedStyles.push(imported);
  }
  for (const paragraph of clonedParagraphs) {
    visitReferences(paragraph, remap);
    const lowered = lowerSourceTheme({ value: paragraph, source });
    if (lowered.isErr()) return lowered;
  }
  const prior = destination.package.styles;
  const styles =
    importedStyles.length === 0
      ? prior
      : {
          ...prior,
          styles: [...(prior?.styles ?? []), ...importedStyles],
        };
  return Result.ok({ styles, paragraphs: clonedParagraphs, styleIds, numberingIds });
};
