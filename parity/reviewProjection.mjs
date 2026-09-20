import { FolioDocxReviewer } from "@stll/folio-core/server";

/** Resolve the accepted package without pulling the server API's large type
 * graph into the parity harness compiler boundary. */
export const projectFinalReviewView = async (source) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  if (reviewer.acceptAll() === 0) {
    return source;
  }
  return await reviewer.toBuffer();
};

/** Test-only structural observation kept at the same isolated boundary. */
export const countReviewChanges = async (source) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  return reviewer.getChanges().length;
};
