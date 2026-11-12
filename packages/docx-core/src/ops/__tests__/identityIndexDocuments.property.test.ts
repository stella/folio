import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { packageIdOccupancy } from "../editorIntentCensus";
import {
  countKeys,
  IDENTITY_SPACES,
  packageIdentityKeyIndex,
  packageIdentityKeys,
  packageParagraphIds,
  packageReservedIdentityKeys,
  packageReservedKeyIndex,
} from "../ids";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";
import { expectOccupancyOf } from "./occupancyProbe";

setDefaultTimeout(propertyTestTimeout(240_000));

/** Every run checks its starting documents, so a few runs show the checks run. */
const COVERAGE_RUNS = 4;

/**
 * The identity counts one process answers, while it edits several documents
 * and returns to their earlier versions in any order, are the counts a
 * process that never saw another document answers.
 */

type Ids = typeof import("../ids");
let instance = 0;
/** The module evaluated again, with none of the state this process holds. */
const freshIds = async (): Promise<Ids> => {
  instance += 1;
  const module: Ids = await import(`../ids?fresh=${instance}`);
  return module;
};

const sorted = (counts: ReadonlyMap<string, number>) =>
  [...counts].toSorted(([left], [right]) => left.localeCompare(right));

const idsIn = (keys: readonly string[], space: string) =>
  new Set(
    keys
      .filter((key) => key.startsWith(`${space}:`))
      .map((key) => Number(key.slice(space.length + 1))),
  );

const step = fc.record({
  document: fc.nat(),
  version: fc.nat(),
  latest: fc.boolean(),
  tracked: fc.boolean(),
  seed: opSeedArbitrary,
});

describe("identity counts across documents", () => {
  test("equal a fresh count through interleaved edits of several documents", async () => {
    // A fresh evaluation holds caches of its own.
    expect((await freshIds()).packageIdentityKeyIndex).not.toBe(
      packageIdentityKeyIndex,
    );
    const law = (counts: { checked: number }) =>
      fc.asyncProperty(
        fc.array(fc.oneof(documentArbitrary, reviewDocumentArbitrary), {
          minLength: 2,
          maxLength: 4,
        }),
        fc.array(step, { minLength: 4, maxLength: 16 }),
        async (documents, steps) => {
          const histories: Document[][] = documents.map((document) => [
            document,
          ]);
          const check = async (document: Document) => {
            const fresh = await freshIds();
            const identities = fresh.packageIdentityKeys(document.package);
            const reserved = fresh.packageReservedIdentityKeys(
              document.package,
            );
            expect(
              sorted(packageIdentityKeyIndex(document.package).counts()),
            ).toEqual(sorted(fresh.countKeys(identities)));
            expect(
              sorted(packageReservedKeyIndex(document.package).counts()),
            ).toEqual(sorted(fresh.countKeys(reserved)));
            // The process's own full counts agree too.
            expect(
              sorted(countKeys(packageIdentityKeys(document.package))),
            ).toEqual(sorted(fresh.countKeys(identities)));
            expect(
              sorted(countKeys(packageReservedIdentityKeys(document.package))),
            ).toEqual(sorted(fresh.countKeys(reserved)));
            const all = [...identities, ...reserved];
            const expected = {
              revisions: idsIn(all, IDENTITY_SPACES.REVISION),
              controls: idsIn(all, IDENTITY_SPACES.CONTROL),
              paragraphs: new Set(
                fresh.packageParagraphIds(document.package).map(fresh.idKey),
              ),
            };
            expectOccupancyOf(packageIdOccupancy(document.package), expected);
            expect(packageParagraphIds(document.package).length).toBe(
              fresh.packageParagraphIds(document.package).length,
            );
            counts.checked += 1;
          };
          for (const document of documents) await check(document);
          for (const {
            document: pick,
            version,
            latest,
            tracked,
            seed,
          } of steps) {
            const index = pick % histories.length;
            const history = histories[index] ?? [];
            // An edit goes on from the latest version, or from an earlier
            // one, as undo and a second session over one file do.
            const base = latest
              ? history.at(-1)
              : history.at(version % history.length);
            if (base === undefined) continue;
            const op = tracked
              ? trackedOpFor(base, seed, history.length)
              : opFor(base, seed);
            const applied = applyDocumentOp(base, op);
            await check(base);
            if (applied.isErr()) continue;
            history.push(applied.value.document);
            await check(applied.value.document);
          }
        },
      );
    await assertProperty(law({ checked: 0 }), {
      numRuns: 40,
      id: "identity counts across documents",
    });
    const counts = { checked: 0 };
    await checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.checked).toBeGreaterThan(0);
  });
});
