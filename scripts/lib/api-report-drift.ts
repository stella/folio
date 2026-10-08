import { existsSync, readFileSync } from "node:fs";
import { renderReportDiff } from "./api-report-diff";

type RenderApiReportDriftOptions = {
  baselinePath: string;
  candidatePath: string;
  reportPath: string;
  maxLines: number;
};

/** A new configured entry has no baseline; print its entire generated report. */
export const renderApiReportDrift = ({
  baselinePath,
  candidatePath,
  reportPath,
  maxLines,
}: RenderApiReportDriftOptions): string => {
  const candidate = readFileSync(candidatePath, "utf8");
  if (!existsSync(baselinePath)) return `new entry: commit ${reportPath}\n${candidate}`;
  return renderReportDiff({
    baseline: readFileSync(baselinePath, "utf8"),
    candidate,
    maxLines,
  });
};
