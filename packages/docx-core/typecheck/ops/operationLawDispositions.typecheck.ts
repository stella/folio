import {
  OPERATION_LAW_DISPOSITIONS,
  type OperationLawDisposition,
} from "../../../../test/operation-law-dispositions";
import type { DocumentOp } from "../../src/ops/types";

export const operationLawDispositions =
  OPERATION_LAW_DISPOSITIONS satisfies Record<
    DocumentOp["type"],
    OperationLawDisposition
  >;
