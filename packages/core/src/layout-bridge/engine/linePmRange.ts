import type { MeasuredLine, ParagraphBlock } from "../../layout-engine/types";

const clampToRunPmEnd = (run: { pmEnd?: number }, position: number) =>
  run.pmEnd === undefined ? position : Math.min(position, run.pmEnd);

export const computeLinePmRange = (
  block: ParagraphBlock,
  line: MeasuredLine,
): { pmStart: number; pmEnd: number } => {
  const fromRun = block.runs[line.fromRun];
  const toRun = block.runs[line.toRun];
  if (fromRun?.pmStart !== undefined && toRun?.pmStart !== undefined) {
    return {
      pmStart:
        fromRun.kind === "text"
          ? clampToRunPmEnd(fromRun, fromRun.pmStart + line.fromChar)
          : fromRun.pmStart,
      pmEnd:
        toRun.kind === "text"
          ? clampToRunPmEnd(toRun, toRun.pmStart + line.toChar)
          : (toRun.pmEnd ?? toRun.pmStart + 1),
    };
  }

  let pmStart = (block.pmStart ?? 0) + 1;
  let lineLength = 0;
  for (
    let runIndex = 0;
    runIndex <= line.toRun && runIndex < block.runs.length;
    runIndex += 1
  ) {
    const run = block.runs[runIndex];
    if (run === undefined) continue;
    if (runIndex < line.fromRun) {
      pmStart += run.kind === "text" ? run.text.length : 1;
      continue;
    }
    if (runIndex === line.fromRun) pmStart += line.fromChar;
    if (run.kind !== "text") {
      lineLength += 1;
      continue;
    }
    const start = runIndex === line.fromRun ? line.fromChar : 0;
    const end = runIndex === line.toRun ? line.toChar : run.text.length;
    lineLength += end - start;
  }
  return { pmStart, pmEnd: pmStart + lineLength };
};
