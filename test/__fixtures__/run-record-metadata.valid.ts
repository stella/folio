type Run = { type: "run"; content: unknown[] };
type Mark = { type: "runIdentity" };
type SourceNode = { type: string; marks: readonly Mark[] };

declare const restoreRunRecord: (run: Run, marks: readonly Mark[]) => void;

const sourceTextBoxAnchor = (node: SourceNode): Run | undefined => {
  if (node.type !== "textBoxAnchor") return undefined;
  const marker: Run = { type: "run", content: [] };
  restoreRunRecord(marker, node.marks);
  return marker;
};

const optionalSourceMarks = (marks?: readonly Mark[]): Run => {
  const run: Run = { type: "run", content: [] };
  if (marks) restoreRunRecord(run, marks);
  return run;
};

const freshGeneratedRun = (): Run => ({
  type: "run",
  content: [],
});

void [sourceTextBoxAnchor, optionalSourceMarks, freshGeneratedRun];
