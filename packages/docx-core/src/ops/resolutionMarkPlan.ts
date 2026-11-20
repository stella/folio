import type { Paragraph } from "../model/document";
import type { DocumentOp, OpStory, ResolveRevisionOp } from "./types";

export type JoinPlan = {
  op: ResolveRevisionOp;
  story: OpStory;
  paraId: string;
  revisionId: number;
  added: boolean;
  depth: number | undefined;
};

export type ResolutionReceiverEffect = {
  type:
    | "paragraph-join"
    | "table-run-in"
    | "block-run-in"
    | "retained-range-boundaries"
    | "section-transfer";
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
};
