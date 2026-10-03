import type { SetSectionPropsOp } from "../../src/ops/types";

// Omission is untouched; null removes a key; explicit undefined owns the key.
const OMITTED = { evenAndOddHeaders: true } as const satisfies SetSectionPropsOp["patch"];
const CLEARED = { titlePg: null } as const satisfies SetSectionPropsOp["patch"];
const OWNED_UNDEFINED = { titlePg: undefined } as const satisfies SetSectionPropsOp["patch"];
// @ts-expect-error a section flag cannot contain a number
const INVALID: SetSectionPropsOp["patch"] = { titlePg: 0 };

export type SectionPatchProof = [
  typeof OMITTED,
  typeof CLEARED,
  typeof OWNED_UNDEFINED,
  typeof INVALID,
];
