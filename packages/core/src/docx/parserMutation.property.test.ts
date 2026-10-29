/**
 * Bounded mutations of small, public-API DOCX fixtures. Each accepted package
 * must save and reopen to the same reader output; a refusal must be typed.
 * The fixed matrix keeps every mutation active, while fast-check shrinks
 * combinations and replays them with PROPERTY_TEST_SEED and its reported path.
 */

import { test } from "bun:test";
import fc from "fast-check";
import assert from "node:assert/strict";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import {
  commentDocument,
  listDocument,
  notesDocument,
  plainDocument,
} from "../../../../test/consumer-scenarios/support/documents";
import {
  assertProperty,
  propertyTestSeed,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  mutateXml,
  PARSER_XML_MUTATIONS,
  type ParserXmlMutation,
} from "./__tests__/parserXmlMutations";
import {
  mutateZip,
  mutateZipSequence,
  ZIP_MUTATIONS,
  type ZipMutation,
} from "./__tests__/parserZipMutations";

const CASE_TIMEOUT_MS = 8_000;
const MAX_MUTATED_BYTES = 512 * 1024;
const BASE_RUNS = 12;
const DEFAULT_SEED = 20_260_927;

const FIXTURES = {
  plain: plainDocument,
  lists: listDocument,
  notes: notesDocument,
  comments: commentDocument,
} as const;

type Fixture = keyof typeof FIXTURES;
type MutationCase = {
  fixture: Fixture;
  part: string;
} & (
  | { layer: "zip"; mutation: ZipMutation }
  | { layer: "xml"; mutation: ParserXmlMutation }
);

const CASES = [
  ...ZIP_MUTATIONS.map((mutation) => ({
    layer: "zip" as const,
    mutation,
    fixture: "plain" as const,
    part: "word/document.xml",
  })),
  ...PARSER_XML_MUTATIONS.map((mutation): MutationCase => {
    switch (mutation) {
      case "outOfRangeId":
        return {
          layer: "xml",
          mutation,
          fixture: "lists",
          part: "word/numbering.xml",
        };
      case "cyclicStyles":
        return {
          layer: "xml",
          mutation,
          fixture: "plain",
          part: "word/styles.xml",
        };
      case "danglingNumberingId":
        return {
          layer: "xml",
          mutation,
          fixture: "lists",
          part: "word/document.xml",
        };
      case "danglingCommentId":
        return {
          layer: "xml",
          mutation,
          fixture: "comments",
          part: "word/document.xml",
        };
      case "danglingFootnoteId":
        return {
          layer: "xml",
          mutation,
          fixture: "notes",
          part: "word/document.xml",
        };
      case "selfReferencingRelationship":
        return {
          layer: "xml",
          mutation,
          fixture: "plain",
          part: "_rels/.rels",
        };
      case "reorderElements":
        return {
          layer: "xml",
          mutation,
          fixture: "plain",
          part: "word/styles.xml",
        };
      default:
        return {
          layer: "xml",
          mutation,
          fixture: "plain",
          part: "word/document.xml",
        };
    }
  }),
  {
    layer: "xml",
    mutation: "dropElement",
    fixture: "notes",
    part: "word/footnotes.xml",
  },
  {
    layer: "xml",
    mutation: "duplicateElement",
    fixture: "notes",
    part: "word/footnotes.xml",
  },
  {
    layer: "xml",
    mutation: "dropElement",
    fixture: "comments",
    part: "word/comments.xml",
  },
  {
    layer: "xml",
    mutation: "duplicateElement",
    fixture: "comments",
    part: "word/comments.xml",
  },
] satisfies MutationCase[];

const fixtureCache = new Map<Fixture, Promise<Uint8Array>>();
const fixtureBytes = (name: Fixture): Promise<Uint8Array> => {
  const cached = fixtureCache.get(name);
  if (cached) return cached;
  const created = FIXTURES[name]();
  fixtureCache.set(name, created);
  return created;
};

const CASE_RUNNER = join(
  dirname(fileURLToPath(import.meta.url)),
  "__tests__/parserMutationCase.ts",
);

const withCaseTimeout = async (
  mutated: Uint8Array,
  label: string,
): Promise<"opened" | "refused"> => {
  const child = Bun.spawn(
    [process.execPath, CASE_RUNNER, Buffer.from(mutated).toString("base64")],
    {
      stdout: "pipe",
      stderr: "pipe",
    },
  );
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    child.kill(9);
  }, CASE_TIMEOUT_MS);
  try {
    const [exitCode, output, diagnostic] = await Promise.all([
      child.exited,
      new Response(child.stdout).text(),
      new Response(child.stderr).text(),
    ]);
    assert.ok(!timedOut, `${label}: exceeded ${CASE_TIMEOUT_MS} ms`);
    assert.equal(
      exitCode,
      0,
      `${label}: ${diagnostic || `case exited ${exitCode}`}`,
    );
    assert.match(
      output,
      /^(?:opened|refused)\n$/u,
      `${label}: missing case result`,
    );
    return output.trim() === "opened" ? "opened" : "refused";
  } finally {
    clearTimeout(timer);
  }
};

const applyMutation = (
  bytes: Uint8Array,
  { layer, mutation, part }: MutationCase,
) =>
  layer === "zip"
    ? mutateZip(bytes, mutation, part)
    : mutateXml(bytes, mutation, part);

const assertSurvives = async (
  mutated: Uint8Array,
  label: string,
): Promise<"opened" | "refused"> => {
  assert.ok(
    mutated.byteLength <= MAX_MUTATED_BYTES,
    `${label}: mutation exceeded input budget`,
  );
  return withCaseTimeout(mutated, label);
};

const checkCase = async (
  entry: MutationCase,
): Promise<"opened" | "refused"> => {
  const label = `${entry.layer}/${entry.mutation} on ${entry.fixture}:${entry.part}`;
  const original = await fixtureBytes(entry.fixture);
  const mutated = await applyMutation(original, entry);
  assert.notDeepEqual(mutated, original, `${label}: mutation made no change`);
  return assertSurvives(mutated, label);
};

test("parser mutation matrix covers every declared mutator", () => {
  assert.deepEqual(
    new Set(
      CASES.filter(({ layer }) => layer === "zip").map(
        ({ mutation }) => mutation,
      ),
    ),
    new Set(ZIP_MUTATIONS),
  );
  assert.deepEqual(
    new Set(
      CASES.filter(({ layer }) => layer === "xml").map(
        ({ mutation }) => mutation,
      ),
    ),
    new Set(PARSER_XML_MUTATIONS),
  );
});

for (const entry of CASES) {
  test(
    `${entry.layer}/${entry.mutation} on ${entry.fixture}:${entry.part} refuses or saves and reopens`,
    async () => {
      await checkCase(entry);
    },
    CASE_TIMEOUT_MS + 1_000,
  );
}

test("a changed but valid package exercises the save/reopen oracle", async () => {
  const original = await fixtureBytes("plain");
  const changed = await mutateXml(original, "emptyText", "word/document.xml");
  assert.equal(await assertSurvives(changed, "empty text"), "opened");
});

// These mutations preserve the part and can be composed on the same fixture.
const COMPOSABLE_XML = [
  "bom",
  "emptyText",
  "hugeText",
  "unknownNamespace",
] as const;
const COMPOSABLE_ZIP = [
  "duplicate-part",
  "bad-crc",
  "wrong-content-type",
] as const;
const composition = fc.oneof(
  fc
    .uniqueArray(fc.constantFrom(...COMPOSABLE_XML), {
      minLength: 1,
      maxLength: 3,
    })
    .map((mutations) => ({ type: "xml" as const, mutations })),
  fc
    .uniqueArray(fc.constantFrom(...COMPOSABLE_ZIP), {
      minLength: 1,
      maxLength: 3,
    })
    .map((mutations) => ({ type: "zip" as const, mutations })),
);

type Composition =
  typeof composition extends fc.Arbitrary<infer Drawn> ? Drawn : never;

/**
 * Locally a fixed sequence. Under CI, or with PROPERTY_TEST_SEED, the shared
 * seed discipline applies, so a per-commit search counts as exploration.
 */
const localSeed = (): { seed?: number } => {
  const ci = process.env["CI"];
  const underCi = ci !== undefined && ci !== "" && ci !== "false" && ci !== "0";
  return underCi || propertyTestSeed() !== undefined
    ? {}
    : { seed: DEFAULT_SEED };
};

const composeMutations = async ({
  type,
  mutations,
}: Composition): Promise<Uint8Array> => {
  const bytes = await fixtureBytes("plain");
  switch (type) {
    case "xml": {
      let mutated = bytes;
      for (const mutation of mutations) {
        mutated = await mutateXml(mutated, mutation, "word/document.xml");
      }
      return mutated;
    }
    case "zip":
      return mutateZipSequence(bytes, mutations, "word/document.xml");
    default: {
      const unhandled: never = type;
      throw new Error(`Unhandled composition ${JSON.stringify(unhandled)}`);
    }
  }
};

/** Every non-empty sequence of distinct items, in every order. */
const sequences = <T>(items: readonly T[]): T[][] => {
  const found: T[][] = [];
  for (const [index, item] of items.entries()) {
    found.push([item]);
    for (const rest of sequences(items.toSpliced(index, 1))) {
      found.push([item].concat(rest));
    }
  }
  return found;
};

test("every order of composable ZIP mutations produces a mutated package", async () => {
  const original = await fixtureBytes("plain");
  for (const mutations of sequences(COMPOSABLE_ZIP)) {
    const mutated = await composeMutations({ type: "zip", mutations });
    assert.notDeepEqual(
      mutated,
      original,
      `zip: ${mutations.join(" + ")} made no change`,
    );
  }
});

test(
  "seeded parser mutation sequences refuse or round-trip",
  async () => {
    await assertProperty(
      fc.asyncProperty(composition, async (entry) => {
        await assertSurvives(
          await composeMutations(entry),
          `${entry.type}: ${entry.mutations.join(" + ")}`,
        );
      }),
      {
        ...localSeed(),
        id: "seeded parser mutation sequences refuse or round-trip",
        numRuns: Number(process.env["PARSER_FUZZ_RUNS"] ?? BASE_RUNS),
      },
    );
  },
  propertyTestTimeout(60_000),
);
