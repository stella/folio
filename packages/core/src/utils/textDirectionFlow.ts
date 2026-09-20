/**
 * The CSS one `w:textDirection` flow renders as.
 *
 * `ST_TextDirection` spells each of its six flows twice — a Strict spelling
 * and a Transitional one — so rendering is decided per flow and read through
 * `TEXT_DIRECTION_FLOW_BY_TOKEN`. Both CSS projections, the editor's inline
 * style and `tableCellToStyle`, read this one map: a cell written `tbRl` and
 * its Strict twin written `rl` cannot render differently in one and the same
 * in the other.
 */

import {
  TEXT_DIRECTION_FLOW_BY_TOKEN,
  type TextDirection,
  type TextDirectionFlow,
} from "@stll/docx-core/model";

export type TextFlowCss = {
  /** `horizontal-tb` needs no declaration; the others are written out. */
  writingMode: "horizontal-tb" | "vertical-lr" | "vertical-rl";
  /**
   * CSS has no writing mode for a bottom-to-top flow, so the nearest one is
   * turned upside down.
   */
  rotateDegrees?: 180;
};

/**
 * Each text flow, and the CSS that renders it.
 *
 * `tb` and `tbV` are the horizontal flows. `rl` and `rlV` run their lines
 * right to left, `lrV` runs them left to right, and `lr` runs them left to
 * right with its characters bottom-to-top, which CSS cannot express: it takes
 * the nearest writing mode turned over.
 *
 * Total over the flows on purpose: a flow nobody mapped would otherwise render
 * as horizontal text with nothing to say it had been dropped.
 */
const TEXT_FLOW_CSS = {
  tb: { writingMode: "horizontal-tb" },
  tbV: { writingMode: "horizontal-tb" },
  rl: { writingMode: "vertical-rl" },
  rlV: { writingMode: "vertical-rl" },
  lrV: { writingMode: "vertical-lr" },
  lr: { writingMode: "vertical-lr", rotateDegrees: 180 },
} as const satisfies Record<TextDirectionFlow, TextFlowCss>;

/** The CSS a `w:textDirection` token renders as, through the flow it names. */
export const textFlowCss = (direction: TextDirection): TextFlowCss =>
  TEXT_FLOW_CSS[TEXT_DIRECTION_FLOW_BY_TOKEN[direction]];
