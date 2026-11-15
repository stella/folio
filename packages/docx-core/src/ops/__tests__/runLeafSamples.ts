import type { RunContent } from "../../model/document";

/**
 * One leaf of every kind a run holds. The record is exhaustive, so a new
 * kind of run content fails to typecheck here until it has a sample.
 */
export const RUN_LEAVES: {
  readonly [Kind in RunContent["type"]]: RunContent & { type: Kind };
} = {
  text: { type: "text", text: "x" },
  tab: { type: "tab" },
  break: { type: "break" },
  symbol: { type: "symbol", font: "Symbol", char: "F020" },
  footnoteRef: { type: "footnoteRef", id: 1 },
  endnoteRef: { type: "endnoteRef", id: 1 },
  noteMarker: { type: "noteMarker", kind: "footnote" },
  fieldChar: { type: "fieldChar", charType: "begin" },
  instrText: { type: "instrText", text: " PAGE " },
  softHyphen: { type: "softHyphen" },
  noBreakHyphen: { type: "noBreakHyphen" },
  renderedPageBreak: { type: "renderedPageBreak" },
  preservedXml: { type: "preservedXml", xml: "<w:ptab/>", text: "" },
  drawing: {
    type: "drawing",
    image: {
      type: "image",
      size: { width: 10, height: 10 },
      wrap: { type: "inline" },
    },
  },
  shape: {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 10, height: 10 },
    },
  },
};

export const RUN_LEAF_SAMPLES: readonly RunContent[] =
  Object.values(RUN_LEAVES);
