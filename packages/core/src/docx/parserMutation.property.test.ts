/**
 * Bounded mutations of small, public-API DOCX fixtures. Each accepted package
 * must save and reopen to the same reader output; a refusal must be typed.
 * The fixed matrix keeps every mutation active, while fast-check shrinks
 * combinations and replays them with PROPERTY_TEST_SEED and its reported path.
 */

import assert from "node:assert/strict";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { test } from "bun:test";
import fc from "fast-check";

import {
  commentDocument,
  listDocument,
  notesDocument,
  plainDocument,
} from "../../../../test/consumer-scenarios/support/documents";
import {
  propertyConfig,
  propertyTestSeed,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  PARSER_XML_MUTATIONS,
  mutateXml,
  type ParserXmlMutation,
} from "./__tests__/parserXmlMutations";
import { ZIP_MUTATIONS, mutateZip, type ZipMutation } from "./__tests__/parserZipMutations";

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
} & ({ layer: "zip"; mutation: ZipMutation } | { layer: "xml"; mutation: ParserXmlMutation });

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
        return { layer: "xml", mutation, fixture: "lists", part: "word/numbering.xml" };
      case "cyclicStyles":
        return { layer: "xml", mutation, fixture: "plain", part: "word/styles.xml" };
      case "danglingNumberingId":
        return { layer: "xml", mutation, fixture: "lists", part: "word/document.xml" };
      case "danglingCommentId":
        return { layer: "xml", mutation, fixture: "comments", part: "word/document.xml" };
      case "danglingFootnoteId":
        return { layer: "xml", mutation, fixture: "notes", part: "word/document.xml" };
      case "selfReferencingRelationship":
        return { layer: "xml", mutation, fixture: "plain", part: "_rels/.rels" };
      case "reorderElements":
        return { layer: "xml", mutation, fixture: "plain", part: "word/styles.xml" };
      default:
        return { layer: "xml", mutation, fixture: "plain", part: "word/document.xml" };
    }
  }),
  { layer: "xml", mutation: "dropElement", fixture: "notes", part: "word/footnotes.xml" },
  { layer: "xml", mutation: "duplicateElement", fixture: "notes", part: "word/footnotes.xml" },
  { layer: "xml", mutation: "dropElement", fixture: "comments", part: "word/comments.xml" },
  { layer: "xml", mutation: "duplicateElement", fixture: "comments", part: "word/comments.xml" },
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
    assert.equal(exitCode, 0, `${label}: ${diagnostic || `case exited ${exitCode}`}`);
    assert.match(output, /^(?:opened|refused)\n$/u, `${label}: missing case result`);
    return output.trim() === "opened" ? "opened" : "refused";
  } finally {
    clearTimeout(timer);
  }
};

const applyMutation = (bytes: Uint8Array, { layer, mutation, part }: MutationCase) =>
  layer === "zip" ? mutateZip(bytes, mutation, part) : mutateXml(bytes, mutation, part);

const assertSurvives = async (
  mutated: Uint8Array,
  label: string,
): Promise<"opened" | "refused"> => {
  assert.ok(mutated.byteLength <= MAX_MUTATED_BYTES, `${label}: mutation exceeded input budget`);
  return withCaseTimeout(mutated, label);
};

const checkCase = async (entry: MutationCase): Promise<"opened" | "refused"> => {
  const label = `${entry.layer}/${entry.mutation} on ${entry.fixture}:${entry.part}`;
  const original = await fixtureBytes(entry.fixture);
  const mutated = await applyMutation(original, entry);
  assert.notDeepEqual(mutated, original, `${label}: mutation made no change`);
  return assertSurvives(mutated, label);
};

test("parser mutation matrix covers every declared mutator", () => {
  assert.deepEqual(
    new Set(CASES.filter(({ layer }) => layer === "zip").map(({ mutation }) => mutation)),
    new Set(ZIP_MUTATIONS),
  );
  assert.deepEqual(
    new Set(CASES.filter(({ layer }) => layer === "xml").map(({ mutation }) => mutation)),
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
const COMPOSABLE_XML = ["bom", "emptyText", "hugeText", "unknownNamespace"] as const;
const COMPOSABLE_ZIP = ["duplicate-part", "bad-crc", "wrong-content-type"] as const;
const composition = fc.oneof(
  fc
    .uniqueArray(fc.constantFrom(...COMPOSABLE_XML), { minLength: 1, maxLength: 3 })
    .map((mutations) => ({ type: "xml" as const, mutations })),
  fc
    .uniqueArray(fc.constantFrom(...COMPOSABLE_ZIP), { minLength: 1, maxLength: 3 })
    .map((mutations) => ({ type: "zip" as const, mutations })),
);

const commitSeed = (): number => {
  const head = process.env["GITHUB_SHA"]?.slice(0, 8);
  return head && /^[0-9a-f]{8}$/iu.test(head) ? Number.parseInt(head, 16) : DEFAULT_SEED;
};

test(
  "seeded parser mutation sequences refuse or round-trip",
  async () => {
    await fc.assert(
      fc.asyncProperty(composition, async ({ type, mutations }) => {
        let bytes = await fixtureBytes("plain");
        if (type === "xml") {
          for (const mutation of mutations) {
            bytes = await mutateXml(bytes, mutation, "word/document.xml");
          }
        } else {
          for (const mutation of mutations) {
            bytes = await mutateZip(bytes, mutation, "word/document.xml");
          }
        }
        await assertSurvives(bytes, `${type}: ${mutations.join(" + ")}`);
      }),
      propertyConfig({
        numRuns: Number(process.env["PARSER_FUZZ_RUNS"] ?? BASE_RUNS),
        seed: propertyTestSeed() ?? commitSeed(),
      }),
    );
  },
  propertyTestTimeout(60_000),
);
