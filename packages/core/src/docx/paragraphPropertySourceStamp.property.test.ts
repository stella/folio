/**
 * A paragraph whose source stamp still matches keeps every fact
 * `paragraphPropertySourceState` reads for it.
 *
 * Caches compare stamps instead of reading every fact again, so a write that
 * changes a fact without renewing the stamp would let a cache reuse a result
 * computed from facts that no longer hold. Sequences of the module's writers
 * run over a few paragraphs, some sharing handles through object spreads;
 * after each step, every paragraph whose stamp matches must read the very
 * same facts.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import {
  assignDocumentParagraphPropertySourceContract,
  assignParagraphIndentationProjection,
  assignParagraphPropertySource,
  cloneParagraphWithPropertySource,
  copyParagraphIndentationProjection,
  copyParagraphPropertyCapture,
  copyParagraphPropertySource,
  createProseParagraphWithPropertySource,
  linkParagraphPropertySourceCandidate,
  paragraphPropertySourceStamp,
  paragraphPropertySourceStampMatches,
  paragraphPropertySourceState,
  restoreParagraphIndentationProjection,
  transferProseParagraphPropertySource,
} from "./paragraphPropertySource";

setDefaultTimeout(propertyTestTimeout(60_000));

const POOL = 5;

const stepArbitrary = fc.record({
  kind: fc.constantFrom(
    "source",
    "copySource",
    "copyCapture",
    "copyProjection",
    "restoreProjection",
    "assignProjection",
    "spread",
    "prose",
    "transfer",
    "candidate",
    "contract",
  ),
  target: fc.nat({ max: POOL - 1 }),
  source: fc.nat({ max: POOL - 1 }),
  indent: fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
  inherited: fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
});
type Step = fc.RecordValue<typeof stepArbitrary>;

const paragraphAt = (index: number, indent: number | undefined): Paragraph => ({
  type: "paragraph",
  paraId: (0x10000000 + index).toString(16).toUpperCase(),
  content: [],
  ...(indent === undefined ? {} : { formatting: { indentLeft: indent * 360 } }),
});

const DIGEST = "ab".repeat(32);

const run = (step: Step, pool: Paragraph[], nodes: PMNode[]): void => {
  const target = pool[step.target] ?? pool[0];
  const source = pool[step.source] ?? pool[0];
  if (target === undefined || source === undefined) return;
  switch (step.kind) {
    case "source":
      assignParagraphPropertySource(
        target,
        `<w:pPr><w:ind w:left="${String(step.indent ?? 0)}"/></w:pPr>`,
      );
      return;
    case "copySource":
      copyParagraphPropertySource(target, source);
      return;
    case "copyCapture":
      copyParagraphPropertyCapture(target, source);
      return;
    case "copyProjection":
      copyParagraphIndentationProjection({ target, source });
      return;
    case "restoreProjection":
      restoreParagraphIndentationProjection({ target, source });
      return;
    case "assignProjection":
      assignParagraphIndentationProjection({
        paragraph: target,
        authored:
          step.indent === undefined
            ? undefined
            : { indentLeft: step.indent * 360 },
        inherited:
          step.inherited === undefined
            ? {}
            : { indentLeft: step.inherited * 360 },
      });
      return;
    // A derivation carries the paragraph's handles to a new object.
    case "spread":
      pool[step.target] = { ...source };
      return;
    case "prose":
      nodes.push(
        createProseParagraphWithPropertySource(
          schema.nodes["paragraph"],
          source,
        ),
      );
      return;
    case "transfer": {
      const from = nodes[step.source % Math.max(1, nodes.length)];
      if (from === undefined) return;
      transferProseParagraphPropertySource(
        schema.nodes["paragraph"].create(),
        from,
        `T${String(step.indent ?? 0)}`,
      );
      return;
    }
    case "candidate": {
      const from = nodes[step.source % Math.max(1, nodes.length)];
      if (from === undefined) return;
      linkParagraphPropertySourceCandidate(target, from);
      return;
    }
    case "contract": {
      const document: Document = {
        package: { document: { content: [...pool] } },
      };
      assignDocumentParagraphPropertySourceContract(document, DIGEST);
      return;
    }
    default:
      step.kind satisfies never;
  }
};

describe("paragraph property source stamps", () => {
  test("a matching stamp means the same kept facts", () => {
    assertProperty(
      fc.property(
        fc.array(
          fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
          {
            minLength: POOL,
            maxLength: POOL,
          },
        ),
        fc.array(stepArbitrary, { minLength: 1, maxLength: 30 }),
        (indents, steps) => {
          const pool = indents.map((indent, index) =>
            paragraphAt(index, indent),
          );
          // Fresh copies of the starting paragraphs, so each case starts clean.
          for (const [index, paragraph] of pool.entries())
            pool[index] = cloneParagraphWithPropertySource(paragraph, {});
          const nodes: PMNode[] = [];
          for (const step of steps) {
            const before = pool.map((paragraph) => ({
              paragraph,
              stamp: paragraphPropertySourceStamp(paragraph),
              state: paragraphPropertySourceState(paragraph),
            }));
            run(step, pool, nodes);
            for (const { paragraph, stamp, state } of before) {
              if (!paragraphPropertySourceStampMatches(paragraph, stamp))
                continue;
              const now = paragraphPropertySourceState(paragraph);
              expect(now.length).toBe(state.length);
              for (const [index, fact] of now.entries())
                expect(fact).toBe(state[index]);
            }
          }
        },
      ),
      {
        id: "paragraph property source stamp matches only unchanged facts",
        numRuns: 500,
      },
    );
  });
});
