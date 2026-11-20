/**
 * Saved-package validity, measured against the package a flow opened. A real
 * document can already break the schema (a `w:zoom` without its percent, a
 * relationship to a part it never shipped); a save that keeps such an issue
 * is not what made the package invalid. Every issue the source does not have,
 * and every extra occurrence of one it has, still fails.
 */
import assert from "node:assert/strict";

import { listDocxPackageIssues } from "@stll/docx-core";

let sourceIssues: readonly string[] = [];

/** Record the issues of the package the current flow opened. */
export const setSourcePackage = async (bytes: Uint8Array): Promise<void> => {
  sourceIssues = await listDocxPackageIssues(bytes);
};

/** The issues a save introduced: its own, beyond those the source has. */
export const introducedIssues = (
  saved: readonly string[],
  source: readonly string[] = sourceIssues,
): string[] => {
  const inherited = new Map<string, number>();
  for (const issue of source)
    inherited.set(issue, (inherited.get(issue) ?? 0) + 1);
  return saved.filter((issue) => {
    const left = inherited.get(issue) ?? 0;
    if (left === 0) return true;
    inherited.set(issue, left - 1);
    return false;
  });
};

export const assertSavedPackageValid = async (
  bytes: Uint8Array | ArrayBuffer,
  context: string,
): Promise<void> => {
  const introduced = introducedIssues(await listDocxPackageIssues(bytes));
  assert.deepEqual(
    introduced,
    [],
    `${context}: saved OOXML invalid: ${introduced.join("; ")}`,
  );
};
