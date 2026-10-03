import type { DocumentOp } from "../../src/ops/types";
import {
  OPERATION_LAW_DISPOSITIONS,
  type OperationLawDisposition,
} from "../../../../test/operation-law-dispositions";

export const operationLawDispositions = OPERATION_LAW_DISPOSITIONS satisfies Record<
  DocumentOp["type"],
  OperationLawDisposition
>;
