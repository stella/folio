import { Result, TaggedError } from "better-result";

export class ExactObjectValidationError extends TaggedError("ExactObjectValidationError")<{
  message: string;
  label: string;
  missing: string[];
  extra: string[];
}> {}

type KeyDifferences = { missing: string[]; extra: string[] };

export const keyDifferences = (
  actual: readonly string[],
  expected: readonly string[],
): KeyDifferences => ({
  missing: expected.filter((key) => !actual.includes(key)),
  extra: actual.filter((key) => !expected.includes(key)),
});

const isUnknownRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

export const validateExactObjectKeys = (
  value: unknown,
  expected: readonly string[],
  label: string,
) => {
  if (!isUnknownRecord(value)) {
    return Result.err(
      new ExactObjectValidationError({
        message: `Invalid ${label}: expected an object`,
        label,
        missing: [...expected],
        extra: [],
      }),
    );
  }
  const { missing, extra } = keyDifferences(Object.keys(value), expected);
  if (missing.length > 0 || extra.length > 0) {
    return Result.err(
      new ExactObjectValidationError({
        message:
          `Invalid ${label} keys; missing: ${missing.join(", ") || "none"}; ` +
          `extra: ${extra.join(", ") || "none"}`,
        label,
        missing,
        extra,
      }),
    );
  }
  return Result.ok(value);
};
