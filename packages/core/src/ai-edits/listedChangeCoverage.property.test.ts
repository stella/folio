/**
 * Every revision the model holds belongs to exactly one change of the list a
 * host reads. Hosts resolve by listed id, so a revision no listed change
 * reaches stays pending after the host resolves everything it sees.
 */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { applyDocumentOp } from "@stll/docx-core/ops";

import { listedChangeCoverageProblem } from "../../../../test/listedChangeCoverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  fixture,
  operations,
  seedArbitrary,
} from "../../typecheck/ops/reviewGenerators.typecheck";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { parseShapeDocument } from "../__tests__/editorHarness";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";

const coverageProblem = (document: Document) =>
  listedChangeCoverageProblem(document, toProseDoc(document));

describe("the change list reaches every revision", () => {
  for (const shape of DOCUMENT_SHAPES)
    test(`shape ${shape.id}`, async () => {
      expect(
        coverageProblem(await parseShapeDocument(await shape.build())),
      ).toBeNull();
    });

  for (const [kind, operation] of Object.entries(operations))
    test(
      `a tracked ${kind}`,
      async () => {
        await assertProperty(
          fc.asyncProperty(seedArbitrary, async (seed) => {
            const document = await fixture(seed);
            const applied = applyDocumentOp(
              document,
              operation({ document, seed }),
            );
            if (applied.isErr()) throw applied.error;
            expect(coverageProblem(applied.value.document)).toBeNull();
          }),
          { numRuns: 20, id: "the change list reaches a tracked operation" },
        );
      },
      propertyTestTimeout(60_000),
    );
});
