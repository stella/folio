/**
 * Every operation a canonical edit compiles to, and every inverse applying it
 * returns, decodes through the batch wire boundary and back unchanged.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import path from "node:path";

import { checkCoverageSample } from "../../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document } from "../../../model/document";
import {
  documentArbitrary,
  reviewDocumentArbitrary,
} from "../../__tests__/documentArbitraries";
import { applyDocumentOps } from "../../apply";
import { createEditorIntentIdAllocator } from "../../editorIntent";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../../types";
import { BATCH_WIRE_OP_TYPES, parseDocumentBatch } from "../envelope";
import {
  canonicalBase,
  compileCanonicalEdit,
  editSeedArbitrary,
} from "./canonicalEdits";

setDefaultTimeout(propertyTestTimeout(120_000));

const NUM_RUNS = 60;

const sendable = (ops: readonly DocumentOp[], observed: Set<string>) => {
  const batch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "coverage",
    actor: "actor",
    baseRev: 0,
    ops,
  };
  for (const op of ops) {
    observed.add(op.type);
    expect(BATCH_WIRE_OP_TYPES).toContain(op.type);
  }
  const decoded = parseDocumentBatch(JSON.stringify(batch));
  if (decoded.isErr())
    throw new Error(
      `${ops.map(({ type }) => type).join(", ")} must be sendable: ${decoded.error.message}`,
    );
  expect(JSON.parse(JSON.stringify(decoded.value))).toEqual(
    JSON.parse(JSON.stringify(batch)),
  );
};

const law = (observed: Set<string>) =>
  fc.property(
    fc.oneof(
      { weight: 3, arbitrary: fc.constant(canonicalBase()) },
      { weight: 1, arbitrary: documentArbitrary },
      { weight: 1, arbitrary: reviewDocumentArbitrary },
    ),
    fc.array(editSeedArbitrary, { minLength: 1, maxLength: 12 }),
    (base: Document, seeds) => {
      const allocate = createEditorIntentIdAllocator();
      let document = base;
      for (const [serial, seed] of seeds.entries()) {
        const author = {
          actor: "author",
          tag: "a",
          lease: undefined,
          allocate,
        };
        const ops = compileCanonicalEdit(document, author, seed, serial);
        if (ops === undefined) continue;
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) continue;
        sendable(ops, observed);
        // Undo sends the inverse with its staleness preconditions intact.
        sendable(applied.value.inverse, observed);
        document = applied.value.document;
      }
    },
  );

describe("canonical operation wire coverage", () => {
  test("every compiled canonical edit and its inverse is sendable", () => {
    assertProperty(law(new Set()), {
      numRuns: NUM_RUNS,
      id: "canonical operation wire coverage",
    });
    const observed = new Set<string>();
    checkCoverageSample(law(observed), NUM_RUNS);
    // The sample reaches text, block, table, list, note, header, package,
    // section and comment operations and the inverses of comment edits.
    for (const type of [
      DOCUMENT_OP_TYPES.INSERT_CONTENT,
      DOCUMENT_OP_TYPES.DELETE_RANGE,
      DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      DOCUMENT_OP_TYPES.INSERT_ROW,
      DOCUMENT_OP_TYPES.DELETE_ROW,
      DOCUMENT_OP_TYPES.INSERT_TABLE,
      DOCUMENT_OP_TYPES.DELETE_TABLE,
      DOCUMENT_OP_TYPES.ADD_NOTE,
      DOCUMENT_OP_TYPES.REMOVE_NOTE,
      DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
      DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
      DOCUMENT_OP_TYPES.CREATE_COMMENT,
      DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
      DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
      DOCUMENT_OP_TYPES.DELETE_COMMENT,
      DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
    ])
      expect(observed).toContain(type);
    // Typed text compiles to content insertion: see the test below.
    expect(observed).not.toContain(DOCUMENT_OP_TYPES.INSERT_TEXT);
    expect(observed.size).toBeGreaterThanOrEqual(30);
  });

  test("no canonical intent or inverse builds a plain text insertion", async () => {
    // `insertText` is a primitive for hosts building operations by hand: the
    // canonical intents compile typed text to `insertContent`, and no inverse
    // is one. No production module of any package constructs it, so the wire
    // coverage sample cannot require it.
    const root = path.join(import.meta.dir, "../../../../../..");
    const builds =
      /\btype:\s*(?:DOCUMENT_OP_TYPES\.INSERT_TEXT\b|"insertText")/u;
    const found: string[] = [];
    let scanned = 0;
    for await (const file of new Bun.Glob("packages/*/src/**/*.{ts,tsx}").scan({
      cwd: root,
    })) {
      if (/(?:\.test\.tsx?$|__tests__|__fixtures__)/u.test(file)) continue;
      scanned += 1;
      if (builds.test(await Bun.file(path.join(root, file)).text()))
        found.push(file);
    }
    expect(scanned).toBeGreaterThan(100);
    expect(found).toEqual([]);
    // The typing intent itself: plain text becomes inserted content.
    const ops = compileCanonicalEdit(
      canonicalBase(),
      {
        actor: "author",
        tag: "a",
        lease: undefined,
        allocate: createEditorIntentIdAllocator(),
      },
      { kind: "type", a: 0, b: 0, suggest: false },
      0,
    );
    const types = ops?.map(({ type }) => type);
    expect(types).toContain(DOCUMENT_OP_TYPES.INSERT_CONTENT);
    expect(types).not.toContain(DOCUMENT_OP_TYPES.INSERT_TEXT);
  });
});
