import { cloneDocumentWithParagraphPropertySources } from "../src/docx/paragraphPropertySource";
import { createStyleResolver } from "../src/prosemirror/styles/styleResolver";
import type {
  ColorValue,
  Document,
  SectionProperties,
  StyleDefinitions,
  TabStop,
  TextFormatting,
} from "../src/types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
  writableCopy,
} from "../src/types/documentModel";

// TypeScript checks `readonly` only on arrays when it compares types, so
// without the model's `WritableRecord` mark a read-only record passes to a
// function that takes the writable record and writes it. Every record type of
// the model carries the mark, so each hand-off below fails to compile; a
// writer copies first, and a reader takes the read-only type.

const writesColor = (color: ColorValue): void => {
  color.rgb = "000000";
};
const writesFormatting = (formatting: TextFormatting): void => {
  formatting.bold = true;
};
const writesTabs = (tabs: TabStop[]): void => {
  tabs.length = 0;
};
const writesEachColor = (colors: readonly ColorValue[]): void => {
  for (const color of colors) color.rgb = "000000";
};
const writesUnderline = (
  underline: NonNullable<TextFormatting["underline"]>,
): void => {
  underline.style = "double";
};
const writesSectionProperties = (properties: SectionProperties): void => {
  properties.marginTop = 0;
};
const writesDocument = (document: Document): void => {
  document.package.document.content.length = 0;
};
const readsColor = (color: DeepReadonly<ColorValue>): string | undefined =>
  color.rgb;

export const plainRecord = (color: DeepReadonly<ColorValue>): void => {
  // @ts-expect-error a read-only color is not the writable color
  writesColor(color);
  writesColor(writableCopy<ColorValue>(color));
  readsColor(color);
};

export const nestedRecord = (
  formatting: DeepReadonly<TextFormatting>,
): void => {
  // @ts-expect-error a read-only formatting is not the writable formatting
  writesFormatting(formatting);
  if (formatting.color) {
    // @ts-expect-error a record read out of a read-only one is read-only too
    writesColor(formatting.color);
  }
  if (formatting.underline) {
    // @ts-expect-error a record holding a read-only record is read-only too
    writesUnderline(formatting.underline);
  }
  writesFormatting(writableCopy<TextFormatting>(formatting));
};

export const readOnlyArray = (
  tabs: readonly DeepReadonly<TabStop>[],
  colors: readonly DeepReadonly<ColorValue>[],
): void => {
  // @ts-expect-error a read-only array is not the writable array
  writesTabs(tabs);
  // @ts-expect-error a read-only array's entries are not writable entries
  writesEachColor(colors);
  writesEachColor(writableCopy<ColorValue[]>(colors));
};

export const readOnlyDocument = (model: ReadonlyDocument): void => {
  // @ts-expect-error the model is not a writable document
  writesDocument(model);
  const properties = model.package.document.finalSectionProperties;
  if (properties) {
    // @ts-expect-error the model's section properties are the model's own
    writesSectionProperties(properties);
  }
  writesDocument(cloneDocumentWithParagraphPropertySources(model));
  // The one sanctioned reader hands a reader the model as it is.
  readModel(model);
};

export const styleResults = (styles: StyleDefinitions): void => {
  const resolver = createStyleResolver(styles);
  const run = resolver.resolveRunStyle("Strong");
  if (run) {
    // @ts-expect-error a resolved run style is the engine's cached result
    writesFormatting(run);
  }
  const defaults = resolver.getDocDefaults()?.rPr;
  if (defaults?.color) {
    // @ts-expect-error the document defaults' color is the package's own
    writesColor(defaults.color);
  }
};

// The mark costs a writer nothing: a writable value still passes where a
// read-only one is expected, and an object literal needs no mark.
export const writableStaysAssignable = (formatting: TextFormatting): void => {
  const readOnly: DeepReadonly<TextFormatting> = formatting;
  void readOnly;
  writesColor({ rgb: "FF0000" });
  writesTabs([{ position: 720, alignment: "left" }]);
};
