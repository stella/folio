type Run = { type: "run"; content: unknown[] };
type Mark = { type: "runIdentity" };
type SourceNode = { type: string; marks: readonly Mark[] };

declare const restoreRunRecord: (run: Run, marks: readonly Mark[]) => void;

const missingRestore = (node: SourceNode): Run => {
  void node.marks;
  if (node.type !== "textBoxAnchor") throw new Error("Expected anchor");
  const marker: Run = { type: "run", content: [] };
  return marker;
};

const restoresWrongRun = (node: SourceNode): Run => {
  const marker: Run = { type: "run", content: [] };
  const other: Run = { type: "run", content: [] };
  restoreRunRecord(other, node.marks);
  return marker;
};

const restoresInAnotherSwitchArm = (
  node: SourceNode,
  state: "build" | "restore",
): Run => {
  let run: Run = missingRestore(node);
  switch (state) {
    case "build": {
      run = { type: "run", content: [] };
      break;
    }
    case "restore": {
      restoreRunRecord(run, node.marks);
      break;
    }
  }
  return run;
};

const shadowedRunRestoration = (marks?: readonly Mark[]): Run => {
  const run: Run = { type: "run", content: [] };
  if (marks) {
    {
      const run: Run = { type: "run", content: [] };
      restoreRunRecord(run, marks);
    }
  }
  return run;
};

const restorationAfterEarlyReturn = (marks: readonly Mark[]): Run => {
  const run: Run = { type: "run", content: [] };
  if (marks.length === 0) return run;
  restoreRunRecord(run, marks);
  return run;
};

const arbitraryMarksPredicate = (node: SourceNode): Run => {
  const run: Run = { type: "run", content: [] };
  if (node.marks && node.type === "textBoxAnchor") {
    restoreRunRecord(run, node.marks);
  }
  return run;
};

const nestedPredicate = (node: SourceNode): Run => {
  const run: Run = { type: "run", content: [] };
  if (node.marks) {
    if (node.type === "textBoxAnchor") {
      restoreRunRecord(run, node.marks);
    }
  }
  return run;
};

void [
  missingRestore,
  restoresWrongRun,
  restoresInAnotherSwitchArm,
  shadowedRunRestoration,
  restorationAfterEarlyReturn,
  arbitraryMarksPredicate,
  nestedPredicate,
];
