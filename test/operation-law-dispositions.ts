import type { DocumentOp } from "../packages/docx-core/src/ops/types";

export type KnownOperationLawIssue = {
  knownIssue: "T1" | "T2" | "T3" | "T4" | "T5" | "T6" | "T7";
  fingerprint: string;
};
export type OperationLawDisposition =
  | "holds"
  | KnownOperationLawIssue
  | readonly KnownOperationLawIssue[];

/** Every operation needs an explicit law disposition; fixed cases return to holds. */
export const OPERATION_LAW_DISPOSITIONS = {
  createHeaderFooter: "holds",
  removeHeaderFooter: "holds",
  addNote: "holds",
  removeNote: "holds",
  setSectionProps: "holds",
  restoreStoryParts: "holds",
  deleteBlocks: "holds",
  insertBlocks: "holds",
  insertText: "holds",
  insertContent: "holds",
  deleteRange: "holds",
  splitInline: "holds",
  joinInline: "holds",
  setRunProps: "holds",
  setParagraphProps: [
    { knownIssue: "T4", fingerprint: "7bda296807167052" },
    { knownIssue: "T4", fingerprint: "95820d436833c3fd" },
  ],
  splitBlock: "holds",
  joinBlocks: [
    { knownIssue: "T4", fingerprint: "20c2d3fdbf109520" },
    { knownIssue: "T4", fingerprint: "af2bdcaabf37424e" },
  ],
  replaceBlocks: "holds",
  setParagraphReview: "holds",
  replaceInline: "holds",
  resolveRevision: "holds",
  insertTable: "holds",
  deleteTable: "holds",
  setContainerBlocks: "holds",
  insertRow: "holds",
  deleteRow: "holds",
  setTableRows: "holds",
} as const satisfies Record<DocumentOp["type"], OperationLawDisposition>;
