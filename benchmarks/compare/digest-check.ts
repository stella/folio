type Digest = { buffer: string; changes: string };

type MeasuredDigest = { id: string; digests: Digest };

/** A check is complete only when every measured configuration has a baseline. */
export const checkRecordedDigests = (
  measured: readonly MeasuredDigest[],
  recorded: Readonly<Record<string, Digest>>,
) => {
  const drifted: string[] = [];
  const missing: string[] = [];
  for (const { id, digests } of measured) {
    const previous = recorded[id];
    if (!previous) {
      missing.push(id);
    } else if (previous.buffer !== digests.buffer || previous.changes !== digests.changes) {
      drifted.push(id);
    }
  }
  return { drifted, missing, checked: measured.length - missing.length };
};
