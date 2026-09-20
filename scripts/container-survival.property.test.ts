/**
 * The survival law over the containers folio's parsers own.
 *
 * `bun scripts/container-survival-census.ts run --check` sweeps the whole
 * schema; this test takes the spine — the paragraph, run, table and section
 * containers every document is made of — so a regression there fails
 * `bun test scripts` rather than waiting for the census. It asserts the same
 * thing the census does, against the same committed baseline: a pair the
 * baseline does not list must survive, and a pair it lists must still be lost
 * by the mechanism recorded, because a loss that quietly changes mechanism is a
 * different defect wearing the old one's name.
 *
 * The second half is the value sweep. The exhaustive pass writes one
 * representative value per attribute; a slot can survive that one and lose
 * every other, so the values an attribute's simple type accepts are drawn from
 * the type itself and each has to come back equal.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../test/property-testing";
import type { SurvivalBaseline } from "./container-survival-census";
import { allSubjects, valueKey } from "./container-survival-census";
import type { Subject } from "./lib/container-survival/fixture";
import { runSurvivalLaws, subjectKey, SURVIVAL_LAWS } from "./lib/container-survival/laws";
import { loadContainerSpace, WML_NAMESPACE } from "./lib/container-survival/schemaSpace";
import { valuesForType } from "./lib/container-survival/values";

setDefaultTimeout(propertyTestTimeout(30_000));

const baseline = (await Bun.file(
  new URL("../specifications/container-contract/survival-baseline.json", import.meta.url),
).json()) as SurvivalBaseline;

const space = await loadContainerSpace();

/**
 * The containers a document is made of.
 *
 * Every one of them is on the parse hot path and owns a family of losses the
 * census found. Drawings, maths and the DrawingML vocabularies are swept by the
 * census alone: they are large, they are slow, and their losses are a different
 * programme.
 */
const SPINE = new Set([
  "body",
  "fldSimple",
  "hyperlink",
  "p",
  "pPr",
  "r",
  "rPr",
  "sectPr",
  "tbl",
  "tblGrid",
  "tblGridChange",
  "tblPr",
  "tc",
  "tcPr",
  "tr",
  "trPr",
]);

const onSpine = (subject: Subject): boolean => {
  const { namespace, name } = subject.slot.container.element;
  return namespace === WML_NAMESPACE && SPINE.has(name);
};

const spineSubjects = allSubjects(space).filter(onSpine);

const describeDisagreement = (
  key: string,
  recorded: string | undefined,
  measured: string | undefined,
): string => {
  if (recorded === undefined) {
    return `newly lost: ${key} (${measured})`;
  }
  if (measured === undefined) {
    return `fixed but not locked in: ${key}; re-run the census with --write-baseline`;
  }
  return `mechanism changed: ${key} (${recorded} -> ${measured})`;
};

describe("the survival law agrees with the committed baseline", () => {
  test("every spine pair is where the baseline says it is", async () => {
    expect(spineSubjects.length).toBeGreaterThan(100);
    const disagreements: string[] = [];
    for (const subject of spineSubjects) {
      const outcome = await runSurvivalLaws(space, subject);
      if (outcome.unrepresentable !== null) {
        continue;
      }
      const recorded = baseline.losses[outcome.key];
      const measured =
        outcome.laws[SURVIVAL_LAWS.parse] === false ? "threw" : (outcome.mechanism ?? undefined);
      if (recorded === measured) {
        continue;
      }
      disagreements.push(describeDisagreement(outcome.key, recorded, measured));
    }
    expect(disagreements).toEqual([]);
  }, 180_000);
});

/**
 * `w:background` is a page backdrop, and the schema lets it hold a `w:drawing`.
 *
 * A chain through it is two steps where the structural one is four, so a
 * shortest-path walk sent every DrawingML container down it and the census
 * measured 118 pairs against a backdrop folio does not model — as
 * `the-container-itself-is-lost`, one loss per pair, for a container no
 * document puts a picture in. The step weighting keeps the chain on the spine,
 * and this is what says so: a fixture realism table is not self-checking, and
 * the symptom of getting it wrong is a plausible-looking loss rather than an
 * error.
 */
describe("the census measures a container where a document puts it", () => {
  test("a backdrop is on nobody's chain but its own", () => {
    const throughBackdrop = [...space.containers.values()]
      .filter(
        (container) =>
          container.id.element.name !== "background" &&
          container.path.some(
            ({ element }) => element.namespace === WML_NAMESPACE && element.name === "background",
          ),
      )
      .map((container) => container.id.element.name);
    expect(throughBackdrop).toEqual([]);
  });

  test("a drawing is measured inside a run", () => {
    const drawing = [...space.containers.values()].find(
      ({ id }) => id.element.namespace === WML_NAMESPACE && id.element.name === "drawing",
    );
    expect(drawing?.path.map(({ element }) => element.name)).toEqual([
      "document",
      "body",
      "p",
      "r",
      "drawing",
    ]);
  });
});

/**
 * Every WordprocessingML attribute the baseline says survives, grouped by its container.
 *
 * The census writes one representative value per attribute; a slot can survive
 * that one and lose every other, which is how an enumeration member nobody
 * mapped and a `0` a truthiness test swallowed both hide. Here the values come
 * from the attribute's own simple type, so a schema refresh that adds a member
 * adds a case without anybody remembering to.
 */
type AttributeSubject = Extract<Subject, { kind: "attribute" }>;

const survivingAttributes = allSubjects(space).filter(
  (subject): subject is AttributeSubject =>
    subject.kind === "attribute" &&
    subject.slot.container.element.namespace === WML_NAMESPACE &&
    baseline.losses[subjectKey(subject)] === undefined &&
    valuesForType(space.index, subject.slot.typeQName).values.length > 1,
);

const byContainer = new Map<string, AttributeSubject[]>();
for (const subject of survivingAttributes) {
  const name = subject.slot.container.element.name;
  byContainer.set(name, [...(byContainer.get(name) ?? []), subject]);
}

describe("a slot that survives its representative value survives the rest of its type", () => {
  for (const [container, subjects] of [...byContainer].sort(([left], [right]) =>
    left.localeCompare(right),
  )) {
    test(`w:${container}`, async () => {
      for (const subject of subjects) {
        const { values } = valuesForType(space.index, subject.slot.typeQName);
        await fc.assert(
          fc.asyncProperty(fc.constantFrom(...values), async (value) => {
            const candidate: Subject = { kind: "attribute", slot: subject.slot, value };
            const outcome = await runSurvivalLaws(space, candidate);
            if (outcome.unrepresentable !== null) {
              return;
            }
            const recorded =
              value === subject.value ? undefined : baseline.valueLosses[valueKey(candidate)];
            const measured =
              outcome.laws[SURVIVAL_LAWS.parse] === false
                ? "threw"
                : (outcome.mechanism ?? undefined);
            const slot = valueKey(candidate).replaceAll(`{${WML_NAMESPACE}}`, "w:");
            expect({ slot, measured }).toEqual({ slot, measured: recorded });
          }),
          propertyConfig({ numRuns: Math.min(values.length, 8) }),
        );
      }
    }, 120_000);
  }
});
