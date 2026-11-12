/**
 * Every operation a canonical edit compiles to, and every inverse applying it
 * returns, decodes through the batch wire boundary and back unchanged.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

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
import {
  BATCH_WIRE_OP_TYPES,
  parseDocumentBatch,
  UNSENT_OP_TYPES,
} from "../envelope";
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
    // The sample reaches text, block, table, list, note, package and
    // section operations, and none of them is an unsent kind.
    for (const type of [
      DOCUMENT_OP_TYPES.INSERT_TEXT,
      DOCUMENT_OP_TYPES.DELETE_RANGE,
      DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      DOCUMENT_OP_TYPES.INSERT_ROW,
      DOCUMENT_OP_TYPES.DELETE_ROW,
      DOCUMENT_OP_TYPES.INSERT_TABLE,
      DOCUMENT_OP_TYPES.DELETE_TABLE,
      DOCUMENT_OP_TYPES.ADD_NOTE,
      DOCUMENT_OP_TYPES.REMOVE_NOTE,
      DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
      DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
    ])
      expect(observed).toContain(type);
    for (const type of observed)
      expect(Object.hasOwn(UNSENT_OP_TYPES, type)).toBe(false);
  });
});
