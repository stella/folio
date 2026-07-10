import type {
  FolioBlockDiff,
  FolioVersionDiff,
  FolioVersionDiffSegment,
} from "@stll/folio-core/server";
import { compareDocxVersions as compareDocxVersionsCore } from "@stll/folio-core/server";

/** Backward-compatible name for a core word-level version-diff segment. */
export type FolioAgentVersionDiffSegment = FolioVersionDiffSegment;

/** Backward-compatible name for a core block-level version diff. */
export type FolioAgentBlockDiff = FolioBlockDiff;

/** Backward-compatible name for a structured core version diff. */
export type FolioAgentVersionDiff = FolioVersionDiff;

/** Compare two `.docx` buffers using folio-core's version comparison semantics. */
export const compareDocxVersions = (
  base: ArrayBuffer,
  revised: ArrayBuffer,
): Promise<FolioAgentVersionDiff> => compareDocxVersionsCore(base, revised);

const ELLIPSIS = "…";
const UNCHANGED_EDGE_CHARS = 30;

const truncateUnchangedRun = (text: string): string => {
  if (text.length <= UNCHANGED_EDGE_CHARS * 2 + ELLIPSIS.length) {
    return text;
  }
  return `${text.slice(0, UNCHANGED_EDGE_CHARS)}${ELLIPSIS}${text.slice(-UNCHANGED_EDGE_CHARS)}`;
};

const renderSegment = (segment: FolioVersionDiffSegment): string => {
  switch (segment.type) {
    case "equal":
      return truncateUnchangedRun(segment.text);
    case "del":
      return `[-${segment.text}-]`;
    case "ins":
      return `{+${segment.text}+}`;
    default:
      return "";
  }
};

const formatChangeLine = (change: FolioBlockDiff): string => {
  if (change.type === "added") {
    return `+ [${change.blockId}] added: ${change.text}`;
  }
  if (change.type === "deleted") {
    return `- [${change.blockId}] deleted: ${change.text}`;
  }
  return `~ [${change.blockId}] modified: ${change.segments.map(renderSegment).join("")}`;
};

/** Render a structured version diff as compact, deterministic model input. */
export const formatVersionDiffForLLM = (diff: FolioAgentVersionDiff): string => {
  const { added, deleted, modified, unchanged } = diff.summaryCounts;
  const header = `Version diff: ${added} added, ${deleted} deleted, ${modified} modified, ${unchanged} unchanged`;
  return [header, ...diff.changes.map(formatChangeLine)].join("\n");
};
