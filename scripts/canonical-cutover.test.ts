import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
} from "../packages/core/src/types/canonicalCapabilities";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  canonicalCutoverDocs,
  inspectCanonicalSources,
  PUBLIC_OPERATION_SOURCE_GAPS,
} from "./lib/canonical-cutover";

setDefaultTimeout(propertyTestTimeout(10_000));

const selectors = Object.keys(CANONICAL_GAP)
  .map((key) => `void CANONICAL_GAP.${key};`)
  .join("\n");
const source = (text: string, file = "packages/react/src/fixture.tsx") => ({
  file,
  source: text,
});
const failuresOf = (text: string, file?: string) =>
  inspectCanonicalSources([source(selectors), source(text, file)]).failures;

describe("canonical cutover guard", () => {
  test("every ledger id has a source site", () => {
    const inspected = inspectCanonicalSources([source(selectors)]);
    expect(inspected.failures).toEqual([]);
    expect([...inspected.sites.keys()].sort()).toEqual(
      Object.keys(CANONICAL_CAPABILITIES).sort(),
    );
  });

  test("every removed ledger reference becomes an orphan", () => {
    for (const key of Object.keys(CANONICAL_GAP)) {
      const inspected = inspectCanonicalSources([
        source(selectors.replace(`void CANONICAL_GAP.${key};`, "")),
      ]);
      expect(inspected.failures).toHaveLength(1);
      expect(inspected.failures.at(0)).toContain("has no source site");
    }
  });

  test("rejects unknown and missing gate/refusal ids", () => {
    for (const fixture of [
      "refuseCanonicalModelEdit('free-text');",
      "handleSessionRefusal('Note story unavailable');",
      "new CanonicalSessionRefusalError({ message: 'refused' });",
      "new CanonicalSessionError({ reason: 'refused', message: 'refused' });",
      'import { CANONICAL_GAP as hidden } from "@stll/folio-core/types/canonicalCapabilities"; void hidden.comments;',
    ])
      expect(failuresOf(fixture).length).toBeGreaterThan(0);
  });

  test("gap-aware error relays cannot discard or replace an error's capability", () => {
    assertProperty(
      fc.property(
        fc.constantFrom("prepared", "result", "address"),
        fc.boolean(),
        (name, bracket) => {
          const relay =
            "const forward = (message: string, ground: CanonicalRefusalGround = { gap: CANONICAL_GAP.dispatch }) => {};";
          const error = bracket ? `${name}["error"]` : `${name}.error`;
          const message = bracket ? `${error}["message"]` : `${error}.message`;
          const ground = bracket ? `${error}["ground"]` : `${error}.ground`;
          expect(
            failuresOf(`${relay} forward(${message}, ${ground});`),
          ).toEqual([]);
          for (const argument of [
            "",
            ", { gap: CANONICAL_GAP.dispatch }",
            ", other.error.ground",
          ])
            expect(
              failuresOf(`${relay} forward(${message}${argument});`),
            ).toContain(
              "packages/react/src/fixture.tsx: canonical error relay must forward the original ground",
            );
        },
      ),
      {
        id: "gap-aware error relays cannot discard or replace an error's capability",
      },
    );
  });

  test("retired engine plugins belong in test support", () => {
    const fixtures = [
      [
        'import { history as pmHistory } from "prosemirror-history";',
        "history",
      ],
      [
        'import { columnResizing } from "prosemirror-tables";',
        "columnResizing",
      ],
      [
        'import { tableEditing as repair } from "prosemirror-tables";',
        "tableEditing",
      ],
    ];
    for (const [code, primitive] of fixtures)
      expect(failuresOf(code)).toContain(
        `packages/react/src/fixture.tsx: ${primitive} belongs in the reference editor`,
      );
    expect(
      failuresOf(
        'import { ParaIdAllocatorExtension } from "../__tests__/reference-editor/ParaIdAllocatorExtension";',
      ),
    ).toContain(
      "packages/react/src/fixture.tsx: reference engine plugins belong in test support",
    );
  });

  test("mutation primitives require their actual capability markers", () => {
    const fixtures = [
      {
        gap: CANONICAL_GAP.suggestionPlugin,
        file: "packages/core/src/prosemirror/plugins/suggestionMode.ts",
        code: "const plugin = {\nappendTransaction() {}\n};",
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
    expect(
      [
        ...new Set([...mutationFixtures.map(({ gap }) => gap), ...publicGaps]),
      ].sort(),
    ).toEqual(mutationSources.map(([id]) => id).sort());
    const sourceGaps = [
      ...new Set([...publicGaps, ...PUBLIC_OPERATION_SOURCE_GAPS]),
    ];
    const markers = sourceGaps
      .map((gap) => `// canonical-gap: ${gap}`)
      .join("\n");
    for (const applier of [
      "applyFolioAIEditOperations",
      "applyFolioAIEditOperationsInPackage",
    ]) {
      const call = `\n${applier}({ undefinedReferences: "refuse" });`;
      for (const removed of sourceGaps) {
        const fixture =
          markers
            .split("\n")
            .filter((line) => line !== `// canonical-gap: ${removed}`)
            .join("\n") + call;
        expect(
          failuresOf(
            fixture,
            "packages/core/src/document-operations-apply.ts",
          ).some((failure) =>
            failure.includes(`${removed} source needs its ledger marker`),
          ),
        ).toBe(true);
      }
      expect(
        failuresOf(
          `${markers}${call}`,
          "packages/core/src/document-operations-apply.ts",
        ),
      ).toEqual([]);
    }
  });

  test("save conversions require the routing capability marker", () => {
    const file = "packages/core/src/controller/fixture.ts";
    const code = "fromProseDoc(state.doc, original);";
    expect(failuresOf(code, file)).toContain(
      `${file}:1: ${CANONICAL_GAP.save} source needs its ledger marker`,
    );
    expect(
      failuresOf(`// canonical-gap: ${CANONICAL_GAP.save}\n${code}`, file),
    ).toEqual([]);
  });

  test("cannot keep a dead source alive with an orphan comment", () => {
    expect(
      failuresOf(
        "// canonical-gap: pm-save-projection\nconst unrelated = true;",
      ).some((failure) => failure.includes("orphan canonical gap marker")),
    ).toBe(true);
  });

  test("canonical history functions are not PM history installations", () => {
    expect(failuresOf("const history = () => true; history();")).toEqual([]);
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
      failuresOf(
        "serializeCanonicalSave({ snapshot: { document: fromProseDoc(pm) } });",
      ),
    ).toContain(
      "packages/react/src/fixture.tsx: canonical serialization requires a controller save snapshot",
    );
    expect(
      failuresOf(
        "const snapshot = editor.captureCanonicalSave(); serializeCanonicalSave({ snapshot });",
      ),
    ).toEqual([]);
    expect(
      failuresOf(
        "const snapshot = session.captureSaveSnapshot(); serializeCanonicalSave({ snapshot });",
      ),
    ).toEqual([]);
  });

  test("documentation derives its ids and sites from the same ledger", () => {
    const inspected = inspectCanonicalSources([source(selectors)]);
    const docs = canonicalCutoverDocs(inspected);
    const ids = docs
      .split("\n")
      .filter(
        (line) =>
          line.startsWith("| ") &&
          !line.startsWith("| Id ") &&
          !line.startsWith("| ---"),
      )
      .map((line) => line.split(" | ").at(0)?.slice(2).trim());
    expect(ids.sort()).toEqual(Object.keys(CANONICAL_CAPABILITIES).sort());
    expect(docs).toContain("packages/react/src/fixture.tsx");
  });
});
