/** Only the internal public-operation compiler may borrow a staged document. */
export const canonicalOperationStagingAccess = Symbol(
  "canonical-operation-staging",
);
