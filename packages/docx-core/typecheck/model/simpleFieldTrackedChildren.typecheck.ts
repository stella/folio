import type {
  SimpleField,
  TrackedRunChange,
  ComplexField,
  Hyperlink,
} from "../../src/model/content";

// Every current revision kind is a typed field child, rather than opaque XML.
export type SimpleFieldRevisionProof = TrackedRunChange extends SimpleField["content"][number]
  ? true
  : never;
export const SIMPLE_FIELD_REVISION_PROOF = true satisfies SimpleFieldRevisionProof;

// A ComplexField is run-only; rich results stay in the recursive paragraph model.
export type ComplexFieldHyperlinkProof = Hyperlink extends ComplexField["fieldResult"][number]
  ? never
  : true;
export const COMPLEX_FIELD_HYPERLINK_PROOF = true satisfies ComplexFieldHyperlinkProof;
