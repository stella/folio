import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
  usesCanonicalSession,
} from "../packages/core/src/types/canonicalCapabilities";
import {
  inspectCanonicalSources,
  checkCanonicalBaseline,
  canonicalCutoverDocs,
} from "./lib/canonical-cutover";

setDefaultTimeout(propertyTestTimeout(10_000));

const selectors = Object.keys(CANONICAL_GAP)
  .map((key) => `usesCanonicalSession(experimentalSession, CANONICAL_GAP.${key});`)
  .join("\n");
const source = (text: string, file = "packages/react/src/fixture.tsx") => ({ file, source: text });
const failuresOf = (text: string, file?: string) =>
  inspectCanonicalSources([source(selectors), source(text, file)]).failures;

describe("canonical cutover guard", () => {
  test("every ledger id has a typed selector with exact session semantics", () => {
    const inspected = inspectCanonicalSources([source(selectors)]);
    expect(inspected.failures).toEqual([]);
    expect([...inspected.sites.keys()].sort()).toEqual(Object.keys(CANONICAL_CAPABILITIES).sort());
    for (const id of Object.values(CANONICAL_GAP)) {
      expect(usesCanonicalSession("canonical", id)).toBe(true);
      expect(usesCanonicalSession(undefined, id)).toBe(false);
    }
  });

  test("every removed ledger reference becomes an orphan", () => {
    for (const key of Object.keys(CANONICAL_GAP)) {
      const inspected = inspectCanonicalSources([
        source(
          selectors.replace(`usesCanonicalSession(experimentalSession, CANONICAL_GAP.${key});`, ""),
        ),
      ]);
      expect(inspected.failures).toHaveLength(1);
      expect(inspected.failures.at(0)).toContain("has no source site");
    }
  });

  test("rejects raw selectors across syntax and Vue templates", () => {
    assertProperty(
      fc.property(
        fc.constantFrom(
          "experimentalSession",
          "props.experimentalSession",
          "deps.getExperimentalSession?.()",
          "sessionRef.current",
        ),
        fc.constantFrom("===", "!==", "==", "!="),
        fc.constantFrom('"canonical"', "'canonical'", "undefined"),
        (selector, operator, value) => {
          const expression = `${selector} ${operator} ${value}`;
          for (const fixture of [
            source(`if (${expression}) refuse();`),
            source(
              `<template><button v-if="${expression.replaceAll('"', "'")}"></button></template>`,
              "packages/vue/src/fixture.vue",
            ),
          ])
            expect(
              inspectCanonicalSources([source(selectors), fixture]).failures.some((failure) =>
                failure.includes("raw session comparison"),
              ),
            ).toBe(true);
        },
      ),
    );
  });

  test("namespace-qualified guards count branches and resolve their ledger ids", () => {
    assertProperty(
      fc.property(
        fc.constantFrom(...Object.entries(CANONICAL_GAP)),
        fc.constantFrom("caps", "canonical", "ledger1"),
        fc.boolean(),
        fc.boolean(),
        ([key, gap], namespace, bracketGuard, bracketGap) => {
          const call = bracketGuard
            ? `${namespace}["usesCanonicalSession"]`
            : `${namespace}.usesCanonicalSession`;
          const tag = bracketGap
            ? `${namespace}["CANONICAL_GAP"]["${key}"]`
            : `${namespace}.CANONICAL_GAP.${key}`;
          const fixture = source(
            `import * as ${namespace} from "@stll/folio-core/types/canonicalCapabilities"; if (${call}(experimentalSession, ${tag})) edit();`,
          );
          const inspected = inspectCanonicalSources([
            source(selectors, "packages/react/src/other.tsx"),
            fixture,
          ]);
          expect(inspected.failures).toEqual([]);
          expect(inspected.branches[fixture.file]).toBe(1);
          expect(inspected.sites.get(gap)?.has(fixture.file)).toBe(true);
          expect(
            checkCanonicalBaseline(inspected.branches, {
              "packages/react/src/other.tsx": Object.keys(CANONICAL_GAP).length,
            }),
          ).toContain(`${fixture.file}: session branches increased from 0 to 1`);
        },
      ),
    );
    for (const call of [
      "caps.usesCanonicalSession(experimentalSession)",
      "caps.usesCanonicalSession(experimentalSession, caps.CANONICAL_GAP.unknown)",
      "caps.usesCanonicalSession(experimentalSession, 'free-text')",
      "new caps.CanonicalSessionRefusalError({ message: 'refused' })",
    ]) {
      expect(
        failuresOf(`import * as caps from "@stll/folio-core/types/canonicalCapabilities"; ${call};`)
          .length,
      ).toBeGreaterThan(0);
    }
  });

  test("rejects unknown and missing gate/refusal ids", () => {
    for (const fixture of [
      "usesCanonicalSession(experimentalSession);",
      'import { usesCanonicalSession as hidden } from "@stll/folio-core/types/canonicalCapabilities"; hidden(experimentalSession, CANONICAL_GAP.comments);',
      "const alias = experimentalSession; if (alias === 'canonical') refuse();",
      "if (experimentalSession) refuse();",
      "refuseCanonicalModelEdit('free-text');",
      "handleSessionRefusal('Note story unavailable');",
      "usesCanonicalSession(experimentalSession, 'free-text');",
      "usesCanonicalSession(experimentalSession, CANONICAL_GAP.unknown);",
      "new CanonicalSessionRefusalError({ message: 'refused' });",
      "new CanonicalSessionError({ reason: 'refused', message: 'refused' });",
    ])
      expect(failuresOf(fixture).length).toBeGreaterThan(0);
  });

  test("mutation primitives require their actual capability markers", () => {
    const fixtures = [
      {
        gap: CANONICAL_GAP.history,
        file: "packages/core/src/prosemirror/history.ts",
        code: 'import { history as pmHistory } from "prosemirror-history";\npmHistory();',
      },
      {
        gap: CANONICAL_GAP.tableGeometry,
        file: "packages/core/src/prosemirror/table.ts",
        code: "columnResizing();",
      },
      {
        gap: CANONICAL_GAP.tableGeometry,
        file: "packages/core/src/prosemirror/table.ts",
        code: "tableEditing();",
      },
      {
        gap: CANONICAL_GAP.suggestionPlugin,
        file: "packages/core/src/prosemirror/plugins/suggestionMode.ts",
        code: "const plugin = {\nappendTransaction() {}\n};",
      },
      {
        gap: CANONICAL_GAP.paragraphIdentity,
        file: "packages/core/src/prosemirror/features/ParaIdAllocatorExtension.ts",
        code: "const plugin = {\nappendTransaction() {}\n};",
      },
      {
        gap: CANONICAL_GAP.paragraphTracker,
        file: "packages/core/src/prosemirror/features/ParagraphChangeTrackerExtension.ts",
        code: "const plugin = {\nappendTransaction: () => null\n};",
      },
      {
        gap: CANONICAL_GAP.aiSnapshots,
        file: "packages/core/src/ai-edits/headless.ts",
        code: "class Reviewer {\ncaptureReviewerState() {}\n}",
      },
      {
        gap: CANONICAL_GAP.publicHeadlessSession,
        file: "packages/core/src/ai-edits/headless.ts",
        code: "class Reviewer {\nstatic fromBuffer() {}\n}",
      },
    ];
    for (const { gap, file, code } of fixtures)
      expect(
        failuresOf(code, file).some((failure) =>
          failure.includes(`${gap} source needs its ledger marker`),
        ),
      ).toBe(true);
    const mutationSources = Object.entries(CANONICAL_CAPABILITIES).filter(
      ([, capability]) => capability.kind === "mutation-source",
    );
    const mutationFixtures = fixtures.filter(
      ({ gap }) => CANONICAL_CAPABILITIES[gap].kind === "mutation-source",
    );
    const publicGaps = mutationSources
      .filter(
        ([id, capability]) =>
          capability.owner === "document-operations" &&
          !mutationFixtures.some(({ gap }) => gap === id),
      )
      .map(([id]) => id);
    expect([...new Set([...mutationFixtures.map(({ gap }) => gap), ...publicGaps])].sort()).toEqual(
      mutationSources.map(([id]) => id).sort(),
    );
    const markers = publicGaps.map((gap) => `// canonical-gap: ${gap}`).join("\n");
    for (const removed of publicGaps) {
      const fixture =
        markers
          .split("\n")
          .filter((line) => line !== `// canonical-gap: ${removed}`)
          .join("\n") + "\napplyFolioAIEditOperations({});";
      expect(
        failuresOf(fixture, "packages/core/src/document-operations.ts").some((failure) =>
          failure.includes(`${removed} source needs its ledger marker`),
        ),
      ).toBe(true);
    }
    expect(
      failuresOf(
        `${markers}\napplyFolioAIEditOperations({});`,
        "packages/core/src/document-operations.ts",
      ),
    ).toEqual([]);
  });

  test("save conversions require the routing capability marker", () => {
    const file = "packages/core/src/controller/fixture.ts";
    const code = "fromProseDoc(state.doc, original);";
    expect(failuresOf(code, file)).toContain(
      `${file}:1: ${CANONICAL_GAP.save} source needs its ledger marker`,
    );
    expect(failuresOf(`// canonical-gap: ${CANONICAL_GAP.save}\n${code}`, file)).toEqual([]);
  });

  test("cannot keep a dead source alive with an orphan comment", () => {
    expect(
      failuresOf("// canonical-gap: pm-save-projection\nconst unrelated = true;").some((failure) =>
        failure.includes("orphan canonical gap marker"),
      ),
    ).toBe(true);
  });

  test("canonical history functions are not PM history installations", () => {
    expect(failuresOf("const history = () => true; history();")).toEqual([]);
  });

  test("per-file counts require every decrease and reject every increase", () => {
    assertProperty(
      fc.property(fc.nat({ max: 1000 }), fc.nat({ max: 1000 }), (before, after) => {
        const failures = checkCanonicalBaseline({ "fixture.ts": after }, { "fixture.ts": before });
        expect(failures.length).toBe(before === after ? 0 : 1);
        if (after > before) expect(failures.at(0)).toContain("increased");
        if (after < before) expect(failures.at(0)).toContain("decrease");
      }),
    );
    expect(checkCanonicalBaseline({ "new.ts": 1 }, {})).toHaveLength(1);
    expect(checkCanonicalBaseline({}, { "removed.ts": 1 })).toHaveLength(1);
  });

  test("canonical save ownership rejects PM-derived signals and manufactured snapshots", () => {
    expect(
      inspectCanonicalSources([
        {
          file: "packages/core/src/docx/canonicalSave.ts",
          source: "const ids = getChangedParagraphIds(state);",
        },
      ]).failures,
    ).toContain(
      "packages/core/src/docx/canonicalSave.ts: canonical serialization cannot derive authority from PM",
    );
    expect(
      failuresOf("serializeCanonicalSave({ snapshot: { document: fromProseDoc(pm) } });"),
    ).toContain(
      "packages/react/src/fixture.tsx: canonical serialization requires a controller save snapshot",
    );
    expect(
      failuresOf(
        "const snapshot = editor.captureCanonicalSave(); serializeCanonicalSave({ snapshot });",
      ),
    ).toEqual([]);
  });

  test("documentation derives its ids and sites from the same ledger", () => {
    const inspected = inspectCanonicalSources([source(selectors)]);
    const docs = canonicalCutoverDocs(inspected.sites);
    const ids = docs
      .split("\n")
      .filter(
        (line) => line.startsWith("| ") && !line.startsWith("| Id ") && !line.startsWith("| ---"),
      )
      .map((line) => line.split(" | ").at(0)?.slice(2).trim());
    expect(ids.sort()).toEqual(Object.keys(CANONICAL_CAPABILITIES).sort());
    expect(docs).toContain("packages/react/src/fixture.tsx");
  });
});
