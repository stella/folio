import type {
  FolioBlockDiff,
  FolioCompareDocxVersionsOptions,
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

/** Backward-compatible name for core comparison options. */
export type FolioAgentCompareDocxVersionsOptions = FolioCompareDocxVersionsOptions;

/** Compare two `.docx` buffers using folio-core's version comparison semantics. */
export const compareDocxVersions = (
  base: ArrayBuffer,
  revised: ArrayBuffer,
  options: FolioAgentCompareDocxVersionsOptions = {},
): Promise<FolioAgentVersionDiff> => compareDocxVersionsCore(base, revised, options);

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
  switch (change.type) {
    case "added":
      return `+ [${change.blockId}] added: ${change.text}`;
    case "deleted":
      return `- [${change.blockId}] deleted: ${change.text}`;
    case "modified":
      return `~ [${change.blockId}] modified: ${change.segments.map(renderSegment).join("")}`;
    case "formatChanged":
      return `~ [${change.blockId}] format changed (${change.changedProperties.join(", ")}): ${truncateUnchangedRun(change.text)}`;
    case "movedFrom":
      return `< [${change.blockId}] moved away (move ${change.moveGroupId}): ${truncateUnchangedRun(change.text)}`;
    case "movedTo":
      return `> [${change.blockId}] moved here (move ${change.moveGroupId}): ${truncateUnchangedRun(change.text)}`;
    default:
      return "";
  }
};

/** Render a structured version diff as compact, deterministic model input. */
export const formatVersionDiffForLLM = (diff: FolioAgentVersionDiff): string => {
  const { added, deleted, modified, formatChanged, moved, metadataChanged, unchanged } =
    diff.summaryCounts;
  const metadataSummary = metadataChanged > 0 ? `, ${metadataChanged} metadata-changed` : "";
  const header = `Version diff: ${added} added, ${deleted} deleted, ${modified} modified, ${formatChanged} format-changed, ${moved} moved${metadataSummary}, ${unchanged} unchanged`;
  const metadataLines = diff.metadataChanges.map(
    ({ property, baseValue, revisedValue }) =>
      `~ metadata.${property}: ${JSON.stringify(baseValue)} -> ${JSON.stringify(revisedValue)}`,
  );
  const privacyLines =
    diff.privacyReport.appliedTransforms.length === 0
      ? []
      : [
          `Privacy transforms: ${diff.privacyReport.appliedTransforms.join(", ")}`,
          `Removed metadata fields: ${diff.privacyReport.removedMetadataProperties.join(", ") || "none"}`,
        ];
  return [header, ...privacyLines, ...diff.changes.map(formatChangeLine), ...metadataLines].join(
    "\n",
  );
};
