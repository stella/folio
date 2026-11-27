import type { Paragraph } from "../model/document";
import type { DocumentOp, OpStory, ResolveRevisionOp } from "./types";

export type JoinPlan = {
  op: ResolveRevisionOp;
  story: OpStory;
  paraId: string;
  revisionId: number;
  added: boolean;
  depth: number | undefined;
  /** Paragraphs standing only for their pending insertions. */
  holders: ReadonlySet<string>;
  /** Paragraphs in the authored table whose receiver was pending when review began. */
  pendingReceiverTableIds?: ReadonlySet<string>;
};

export type ResolutionReceiverEffect = {
  type:
    | "paragraph-join"
    | "table-run-in"
    | "block-run-in"
    | "retained-range-boundaries";
  receiver: Paragraph;
};

export type ParagraphMarkPlan = {
  ops: readonly DocumentOp[];
  receiverEffects: readonly ResolutionReceiverEffect[];
  /**
   * The paragraph, ending a block content control, whose join takes an
   * empty paragraph after the control once every other join is done.
   */
  endsControl?: string;
  /**
   * The paragraph, ending a block content control, that the rejection of its
   * added mark left standing only for its pending insertions.
   */
  holdsInsertions?: string;
};
