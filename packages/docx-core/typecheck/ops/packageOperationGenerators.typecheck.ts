import type { DocumentOp } from "../../src/ops/types";
import type { PACKAGE_OP_CASES } from "../../../../test/generators/packageOperationArbitraries";

type Assert<Condition extends true> = Condition;

/** The compiled proof follows the test factory map so no op kind can lack a generator. */
export type PackageOperationGeneratorProof = [
  Assert<Exclude<DocumentOp["type"], keyof typeof PACKAGE_OP_CASES> extends never ? true : false>,
  Assert<Exclude<keyof typeof PACKAGE_OP_CASES, DocumentOp["type"]> extends never ? true : false>,
];
