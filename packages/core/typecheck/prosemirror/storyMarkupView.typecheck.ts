import type { Node as PMNode } from "prosemirror-model";
import { projectStoryMarkupView } from "../../src/prosemirror/markupViewProjection";
import type { convertNoteStoryToFlowBlocks } from "../../src/layout-bridge/convert/footnoteLayout";
import type { convertHeaderFooterPmDocToContent } from "../../src/layout-bridge/convert/headerFooterLayout";

declare const doc: PMNode;
declare const noteOptions: Parameters<typeof convertNoteStoryToFlowBlocks>[1];
declare const headerFooterOptions: Parameters<typeof convertHeaderFooterPmDocToContent>[3];

// Detached renderers pass optional resources explicitly, including absent ones.
projectStoryMarkupView({
  doc,
  markupView: "original",
  styles: noteOptions.styles,
  numbering: noteOptions.numbering,
});
projectStoryMarkupView({
  doc,
  markupView: "no-markup",
  styles: headerFooterOptions.styles,
  numbering: headerFooterOptions.numbering,
});
projectStoryMarkupView({ doc, markupView: "no-markup", styles: undefined, numbering: undefined });
projectStoryMarkupView({ doc, markupView: "all-markup" });
