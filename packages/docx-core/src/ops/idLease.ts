import type { DocumentIdLease } from "../model/identity";

/** A compiler names the paragraph range it delegates to an operation. */
export const leaseForParagraphIds = (ids: readonly string[]): DocumentIdLease => {
  const values = ids.map((id) => Number.parseInt(id, 16));
  return {
    paragraph: {
      start: values.length === 0 ? 1 : Math.max(1, Math.min(...values)),
      end: values.length === 0 ? 1 : Math.max(1, Math.max(...values)),
    },
    revision: { start: 1, end: 0x7fffffff },
    comment: { start: 1, end: 0x7fffffff },
    table: { prefix: "compiler", start: 0, end: Number.MAX_SAFE_INTEGER },
    row: { prefix: "compiler", start: 0, end: Number.MAX_SAFE_INTEGER },
    cell: { prefix: "compiler", start: 0, end: Number.MAX_SAFE_INTEGER },
  };
};
