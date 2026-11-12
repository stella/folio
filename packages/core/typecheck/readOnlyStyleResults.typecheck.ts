import type { EditorState } from "prosemirror-state";

import { getParagraphTabs } from "../src/prosemirror/extensions/core/ParagraphExtension";
import type { SelectionSnapshot } from "../src/prosemirror/selectionState";
import { createStyleResolver } from "../src/prosemirror/styles/styleResolver";
import { createStyleEngine } from "../src/style-engine";
import type {
  ParagraphFormatting,
  StyleDefinitions,
  TextFormatting,
} from "../src/types/document";
import { writableCopy } from "../src/types/documentModel";
import { mergeParagraphFormatting } from "../src/utils/paragraphFormattingMerge";
import { mergeTextFormatting } from "../src/utils/textFormattingMerge";

// The style resolver hands out the document's own style values and results the
// engine caches for every caller, so each one is read-only all the way down: a
// write through any of them fails to compile. A writer copies first.

export const resolverResults = (styles: StyleDefinitions): void => {
  const resolver = createStyleResolver(styles);
  const style = resolver.getStyle("Heading1");
  if (style?.rPr) {
    // @ts-expect-error a style is the document's own definition
    style.rPr.bold = true;
  }
  const run = resolver.resolveRunStyle("Strong");
  if (run?.underline) {
    // @ts-expect-error a resolved run style shares its underline with the document
    run.underline.style = "double";
  }
  const own = resolver.getRunStyleOwnProperties("Strong");
  if (own) {
    // @ts-expect-error a style's own properties are the style's record itself
    own.italic = true;
  }
  const paragraph = resolver.resolveParagraphStyle("Heading1");
  // @ts-expect-error a resolved paragraph style is read-only
  paragraph.paragraphFormatting = {};
  // @ts-expect-error resolved tab stops are the style's own
  paragraph.paragraphFormatting?.tabs?.push({ position: 0, alignment: "left" });
  // The cascade of a style-less paragraph is the document defaults' own
  // record, not a copy: a write to it would change the package.
  const fromDefaults = resolver.resolveParagraphStyle(null).paragraphFormatting;
  if (fromDefaults) {
    // @ts-expect-error the defaults' paragraph properties are the package's own
    fromDefaults.spaceAfter = 0;
  }
  const defaults = resolver.getDocDefaults();
  if (defaults?.rPr) {
    // @ts-expect-error document defaults are the package's own
    defaults.rPr.fontSize = 24;
  }
  const [first] = resolver.getParagraphStyles();
  if (first) {
    // @ts-expect-error a listed style is the document's own definition
    first.name = "Renamed";
  }
};

export const engineResults = (styles: StyleDefinitions): void => {
  const engine = createStyleEngine(styles);
  const paragraph = engine.resolveParagraphStyle(undefined);
  if (paragraph.runFormatting) {
    // @ts-expect-error a cached result is shared with every later caller
    paragraph.runFormatting.bold = false;
  }
  const run = engine.resolveRunStyle(undefined);
  if (run?.color) {
    // @ts-expect-error a cached run style shares its colour with the document
    run.color.rgb = "000000";
  }
  const table = engine.getDefaultTableStyle();
  // @ts-expect-error a table style's conditional regions are the document's own
  table?.tblStylePr?.pop();
};

export const paragraphTabs = (state: EditorState): void => {
  // @ts-expect-error the tab stops are the paragraph's attrs, shared with the cascade
  getParagraphTabs(state)?.push({ position: 720, alignment: "left" });
};

export const selectionFormatting = (snapshot: SelectionSnapshot): void => {
  // @ts-expect-error the selection's run formatting is the editor state's own
  snapshot.textFormatting.bold = true;
  // @ts-expect-error the selection's tab stops are the paragraph's own
  snapshot.paragraphFormatting.tabs?.pop();
};

export const writersCopyFirst = (styles: StyleDefinitions): void => {
  const resolver = createStyleResolver(styles);
  const resolved = resolver.resolveParagraphStyle("Heading1");
  const copy = writableCopy<ParagraphFormatting>(
    resolved.paragraphFormatting ?? {},
  );
  copy.tabs?.push({ position: 720, alignment: "left" });
  const run = writableCopy<TextFormatting>(
    resolver.resolveRunStyle(null) ?? {},
  );
  run.bold = true;
  // A merge of plain formatting stays plain; a merge of resolved formatting
  // is read-only.
  const plain = mergeTextFormatting({ bold: true }, { italic: true });
  if (plain) plain.bold = false;
  const inherited = mergeParagraphFormatting(resolved.paragraphFormatting, {
    spaceAfter: 0,
  });
  inherited.spaceAfter = 120;
  // @ts-expect-error the merge shares its nested values with the style
  inherited.tabs?.push({ position: 720, alignment: "left" });
};
