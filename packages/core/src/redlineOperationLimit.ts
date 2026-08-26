import { TaggedError } from "better-result";

export const MAX_GENERATED_REDLINE_OPERATIONS = 10_000;

export class GenerateRedlineDocxOperationLimitError extends TaggedError(
  "GenerateRedlineDocxOperationLimitError",
)<{
  message: string;
}> {}
