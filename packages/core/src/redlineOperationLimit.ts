import { TaggedError, type TaggedErrorClass } from "better-result";

export const MAX_GENERATED_REDLINE_OPERATIONS = 10_000;

const GenerateRedlineDocxOperationLimitErrorBase: TaggedErrorClass<"GenerateRedlineDocxOperationLimitError"> =
  TaggedError("GenerateRedlineDocxOperationLimitError");

export class GenerateRedlineDocxOperationLimitError extends GenerateRedlineDocxOperationLimitErrorBase<{
  message: string;
}> {}
