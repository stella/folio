import type { SetSectionPropsOp } from "../../src/ops/types";

// Omission leaves a property untouched; null clears it. Undefined is not a wire value.
const OMITTED = { evenAndOddHeaders: true } as const satisfies SetSectionPropsOp["patch"];
const CLEARED = { titlePg: null } as const satisfies SetSectionPropsOp["patch"];
// @ts-expect-error undefined cannot cross the JSON operation boundary
const UNDEFINED: SetSectionPropsOp["patch"] = { titlePg: undefined };

export type SectionPatchProof = [typeof OMITTED, typeof CLEARED, typeof UNDEFINED];
