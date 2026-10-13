import assert from "node:assert/strict";
import { createDocx } from "@stll/folio-core/server";
import { openReviewer } from "./documents.ts";
import { projectContentPair } from "./identity.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type ReviewSource = Pick<Reviewer, "getContent" | "getChanges" | "toDocument">;
type Change = ReturnType<Reviewer["getChanges"]>[number];
type Coverage = Pick<Change, "type" | "author" | "date" | "text" | "blockId">;

/** Adjacent fragments in each provenance stream keep lossless character order. */
export const revisionCoverage = (
  changes: readonly Change[],
  aliases: ReadonlyMap<string, string> = new Map(),
): Coverage[] => {
  const streams = new Map<string, Coverage[]>();
  for (const { type, author, date, text, blockId } of changes) {
    const block = blockId === null ? null : (aliases.get(blockId) ?? blockId);
    const key = JSON.stringify([type, author, date, block]);
    const stream = streams.get(key) ?? [];
    const previous = stream.at(-1);
    if (text.length > 0 && previous && previous.text.length > 0) {
      previous.text += text;
    } else {
      stream.push({ type, author, date, text, blockId: block });
    }
    streams.set(key, stream);
  }
  return [...streams.keys()].sort().flatMap((key) => streams.get(key) ?? []);
};

export const assertRevisionCoverage = (left: readonly Change[], right: readonly Change[]) =>
  assert.deepEqual(revisionCoverage(right), revisionCoverage(left));

const resolved = async (reviewer: ReviewSource, decision: "accept" | "reject") => {
  // Copy the live model independently of the package-saving path under test.
  const copy = await openReviewer(new Uint8Array(await createDocx(reviewer.toDocument())));
  if (decision === "accept") copy.acceptAll();
  else copy.rejectAll();
  const positional = new Set(
    reviewer
      .getContent()
      .filter(({ idStability }) => idStability === "positional")
      .map(({ id }) => id),
  );
  // Compare persisted resolved documents, rather than editor capture metadata.
  const persisted = await openReviewer(new Uint8Array(await copy.toBuffer()));
  const packageModel = persisted.toDocument().package;
  const properties = { ...packageModel.properties };
  // Package save time is bookkeeping; revision authorship dates stay strict above.
  delete properties.modified;
  return {
    package: { ...packageModel, properties },
    rows: persisted
      .getContent()
      .map((row) =>
        positional.has(row.id)
          ? Object.assign({}, row, { idStability: "positional" as const })
          : row,
      ),
  };
};

/** Change only paragraph identities that the shared identity comparator explicitly pairs. */
const packageWithIdentities = (value: unknown, aliases: ReadonlyMap<string, string>): unknown => {
  if (Array.isArray(value)) return value.map((entry) => packageWithIdentities(entry, aliases));
  if (value === null || typeof value !== "object") return value;
  if (Object.getPrototypeOf(value) !== Object.prototype) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [
      key,
      key === "paraId" && typeof entry === "string"
        ? (aliases.get(entry) ?? entry)
        : packageWithIdentities(entry, aliases),
    ]),
  );
};

type ReviewRoundTripOptions = { live: ReviewSource; reopened: ReviewSource; context: string };

/** The save invariant and reader stability share coverage and both resolution outcomes. */
export const assertReviewRoundTrip = async ({
  live,
  reopened,
  context,
}: ReviewRoundTripOptions): Promise<void> => {
  const identities = projectContentPair({
    leftRows: live.getContent(),
    rightRows: reopened.getContent(),
  });
  const leftChanges = live.getChanges();
  const rightChanges = reopened.getChanges();
  assert.deepEqual(
    revisionCoverage(rightChanges, identities.rightAliases),
    revisionCoverage(leftChanges, identities.leftAliases),
    `${context}: revision coverage differs after save`,
  );
  if (leftChanges.length === 0 && rightChanges.length === 0) return;
  for (const decision of ["accept", "reject"] as const) {
    const left = await resolved(live, decision);
    const right = await resolved(reopened, decision);
    const pair = projectContentPair({ leftRows: left.rows, rightRows: right.rows });
    assert.deepEqual(
      packageWithIdentities(right.package, pair.rightAliases),
      packageWithIdentities(left.package, pair.leftAliases),
      `${context}: ${decision}-all document differs after save`,
    );
  }
};
