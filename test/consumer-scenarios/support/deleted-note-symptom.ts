import { AssertionError } from "node:assert";
import { isDeepStrictEqual } from "node:util";

const hasNotes = (value: unknown): value is Record<string, unknown> & { notes: string } =>
  typeof value === "object" &&
  value !== null &&
  "notes" in value &&
  typeof value.notes === "string";

/** Match the compared values, independent of Node's version-specific assertion rendering. */
export const deletedNoteSymptom = (failure: unknown): boolean => {
  const seen = new Set<unknown>();
  let current = failure;
  while (current instanceof Error && !seen.has(current)) {
    seen.add(current);
    if (
      current instanceof AssertionError &&
      current.operator === "deepStrictEqual" &&
      current.message.includes(
        "the reopened package shows something else than the reviewer that saved it",
      ) &&
      hasNotes(current.actual) &&
      hasNotes(current.expected)
    ) {
      const { notes: actualNotes, ...actual } = current.actual;
      const { notes: expectedNotes, ...expected } = current.expected;
      if (!isDeepStrictEqual(actual, expected)) return false;
      const remaining = actualNotes === "" ? [] : actualNotes.split("\n");
      let missing = 0;
      for (const line of expectedNotes.split("\n")) {
        if (line === remaining.at(0)) {
          remaining.shift();
          continue;
        }
        if (!/^\[(?:footnote|endnote) #\d+\] /u.test(line)) return false;
        missing += 1;
      }
      return missing > 0 && remaining.length === 0;
    }
    current = current.cause;
  }
  return false;
};
