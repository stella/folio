/** JSON shared by the corpus runner and its isolated edit worker. */
export type EditStep = {
  mode: "direct" | "tracked-changes";
  operation: { type: string } & Record<string, unknown>;
};

export type EditFailure = {
  class: string;
  signature: string;
  expected: string;
  observed: string;
  operations: EditStep[];
};

export type EditWorkerResult = {
  status: "parsed" | "unparsed";
  attempts: number;
  failures: EditFailure[];
};
