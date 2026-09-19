/**
 * Every part folio writes must still satisfy the schema it claims to write.
 *
 * The invariant is a difference, not an absolute. Public corpora are full of
 * packages that were already invalid when Word accepted them, and folio copies
 * the parts it did not touch through unchanged, so validating the output alone
 * would report the producer's defects as folio's. The check therefore validates
 * the input and the output and reports only what the save introduced: a
 * violation present in a rebuilt part and absent from the same part on the way
 * in is folio writing markup the schema does not allow.
 *
 * The schema is the committed graph under `specifications/generated`, the same
 * one the Strict value-encoding table is generated from, so the two can never
 * disagree about what Transitional means.
 */

import { repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import { Result } from "better-result";

import {
  loadSchemaGraph,
  type SchemaViolation,
  validateOoxmlPart,
} from "../corpus-schema-validator";
import { failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";

/** Violations reported per part, so one pathological part cannot flood the census. */
const VIOLATION_LIMIT = 25;
/** Distinct introduced violations reported per file, for the same reason. */
const REPORTED_PER_FILE = 5;

const DIGITS_RE = /\d+/gu;

/**
 * A part name with its ordinal erased.
 *
 * Losing validity in `word/header2.xml` and in `word/header7.xml` is one
 * defect, and the file that shows it is in the census example.
 */
export const generalizePartPath = (path: string): string => path.replaceAll(DIGITS_RE, "N");

/** Parts whose markup this graph describes. Media, relationships and theme are not WordprocessingML. */
const VALIDATED_PART_RE = /^word\/[\w/]+\.xml$/u;

const violationKey = ({ kind, path, name, detail }: SchemaViolation): string =>
  `${kind} | ${path} | ${name} | ${detail}`;

type PartViolations = Map<string, Set<string>>;

const validatePackage = async (
  bytes: ArrayBuffer,
  graph: Awaited<ReturnType<typeof loadSchemaGraph>>,
): Promise<PartViolations> => {
  const { allXml } = await unzipDocx(bytes, { extractAllXml: true });
  const byPart: PartViolations = new Map();
  for (const [path, xml] of allXml) {
    if (!VALIDATED_PART_RE.test(path)) {
      continue;
    }
    // Keyed by the real path, not the generalised one: two headers generalise
    // to the same name, and collapsing them here would let the last one in zip
    // order hide whatever folio did to the others.
    const violations = validateOoxmlPart({ graph, xml, limit: VIOLATION_LIMIT });
    byPart.set(path, new Set(violations.map(violationKey)));
  }
  return byPart;
};

/**
 * Violations the save introduced.
 *
 * A part the input did not have is compared against nothing, which is correct:
 * folio wrote all of it, so everything wrong with it is folio's.
 */
const introducedViolations = (before: PartViolations, after: PartViolations): string[] => {
  const introduced = new Set<string>();
  for (const [part, violations] of after) {
    const known = before.get(part) ?? new Set<string>();
    for (const violation of violations) {
      if (!known.has(violation)) {
        // Generalised only here: the same defect in header one and header seven
        // is one finding, but the comparison above had to keep them apart.
        introduced.add(`${generalizePartPath(part)} gained ${violation}`);
      }
    }
  }
  return [...introduced].sort();
};

export const runSchemaValidityInvariant = async ({
  buffer,
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};
  const graph = await timeStage(timings, "load-schema", () => loadSchemaGraph());

  const before = await timeStage(timings, "validate-input", () =>
    Result.tryPromise({ try: () => validatePackage(buffer, graph), catch: (c: unknown) => c }),
  );
  if (before.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.schemaValidity, before.error)],
      timings,
    };
  }

  const saved = await timeStage(timings, "save", () =>
    Result.tryPromise({
      try: () => repackDocx(parsed, { updateModifiedDate: false }),
      catch: (cause: unknown) => cause,
    }),
  );
  if (saved.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.schemaValidity, saved.error)],
      timings,
    };
  }

  const after = await timeStage(timings, "validate-output", () =>
    Result.tryPromise({
      try: () => validatePackage(saved.value, graph),
      catch: (cause: unknown) => cause,
    }),
  );
  if (after.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.schemaValidity, after.error)],
      timings,
    };
  }

  const introduced = await timeStage(timings, "compare", () =>
    introducedViolations(before.value, after.value),
  );
  return {
    failures: introduced
      .slice(0, REPORTED_PER_FILE)
      .map((detail) => failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.schemaValidity, detail)),
    timings,
  };
};
