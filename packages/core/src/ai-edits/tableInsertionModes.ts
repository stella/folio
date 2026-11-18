import type { FolioAIEditApplyMode, FolioAIEditOperation } from "./types";

/** The published modes shared by table placement and both execution engines. */
export const TABLE_INSERTION_MODES = Object.freeze({
  insertTable: Object.freeze(["direct", "tracked-changes"] as const),
  insertSignatureTable: Object.freeze(["direct", "suggested"] as const),
} as const satisfies Record<
  Extract<
    FolioAIEditOperation,
    { type: "insertTable" | "insertSignatureTable" }
  >["type"],
  readonly FolioAIEditApplyMode[]
>);

export type TableInsertionOperation = Extract<
  FolioAIEditOperation,
  { type: keyof typeof TABLE_INSERTION_MODES }
>;

export const supportsTableInsertionMode = (
  type: TableInsertionOperation["type"],
  mode: FolioAIEditApplyMode,
): boolean =>
  TABLE_INSERTION_MODES[type].some((supported) => supported === mode);
