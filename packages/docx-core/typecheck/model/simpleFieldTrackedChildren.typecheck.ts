import type { SimpleField, TrackedRunChange } from "../../src/model/content";

// Every current revision kind is a typed field child, rather than opaque XML.
export type SimpleFieldRevisionProof = TrackedRunChange extends SimpleField["content"][number]
  ? true
  : never;
export const SIMPLE_FIELD_REVISION_PROOF = true satisfies SimpleFieldRevisionProof;
