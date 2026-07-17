type DocumentGrid = {
  type?: "default" | "lines" | "linesAndChars" | "snapToChars";
  linePitch?: number;
};

export function resolveDocumentGridLinePitch(grid: DocumentGrid | undefined): number | undefined {
  if (grid?.type === undefined || grid.type === "default") {
    return undefined;
  }
  if (grid.linePitch === undefined || grid.linePitch <= 0) {
    return undefined;
  }
  return grid.linePitch;
}
