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
    ] as const;
    for (const [code, primitive] of fixtures)
      expect(failuresOf(code)).toContain(
        `packages/react/src/fixture.tsx: ${primitive} belongs in the reference editor`,
      );
    expect(
      failuresOf(
        'import * as tables from "prosemirror-tables"; tables["tableEditing"]();',
      ),
    ).toContain(
      "packages/react/src/fixture.tsx: tableEditing belongs in the reference editor",
    );
    expect(
      failuresOf(
        'import * as tables from "prosemirror-tables"; tables.TableMap;',
      ).filter((failure) =>
        failure.includes("belongs in the reference editor"),
      ),
    ).toEqual([]);
    expect(
      failuresOf(
        'import { ParaIdAllocatorExtension } from "../__tests__/reference-editor/ParaIdAllocatorExtension";',
      ),
    ).toContain(
      "packages/react/src/fixture.tsx: reference engine plugins belong in test support",
    );
    const command =
      "packages/core/src/prosemirror/commands/replaceRangeText.ts";
    expect(
      failuresOf(
        'import { deleteSelectionAsSuggestion } from "../../__tests__/reference-editor/suggestionMode";',
        command,
      ),
    ).toContain(`${command}: reference engine plugins belong in test support`);
  });

  test("the headless session factory requires its capability marker", () => {
    const gap = CANONICAL_GAP.publicHeadlessSession;
    const file = "packages/core/src/ai-edits/headless.ts";
    const method = "static fromBuffer() {}";
    expect(failuresOf(`class Reviewer {\n${method}\n}`, file)).toContain(
      `${file}:2: ${gap} source needs its ledger marker`,
    );
    expect(
      failuresOf(
        `class Reviewer {\n// canonical-gap: ${gap}\n${method}\n}`,
        file,
      ),
    ).toEqual([]);
  });

  test("production callers cannot reconstruct a document from PM", () => {
    const file = "packages/core/src/controller/fixture.ts";
    const failure = "the reference editor owns PM document reconstruction";
    expect(failuresOf("fromProseDoc(state.doc, original);", file)).toContain(
      `${file}:1: ${failure}`,
    );
    for (const fixture of [
      'import { fromProseDoc as project } from "../prosemirror/conversion";\nproject(state.doc, original);',
      'import * as conversion from "../prosemirror/conversion";\nconversion["fromProseDoc"](state.doc, original);',
    ])
      expect(failuresOf(fixture, file)).toContain(`${file}:2: ${failure}`);
  });

  test("cannot keep a dead source alive with an orphan comment", () => {
    expect(
      failuresOf(
        `// canonical-gap: ${CANONICAL_GAP.dispatch}\nconst unrelated = true;`,
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
