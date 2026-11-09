/**
 * A paragraph whose facts a reader kept never reads other facts.
 *
 * Caches that convert or normalize a paragraph keep the result by the
 * paragraph's identity and seal its facts when they keep it, so a later write
 * of a fact `paragraphPropertySourceState` reads would let them reuse a
 * result computed from facts that no longer hold. Sequences of the module's
 * writers run over a few paragraphs, some sharing handles through object
 * spreads and some sealed on the way; after each step, whether the writer
 * finished or refused, every sealed paragraph must read the very same facts,
 * and a refusal must name the fact it would have changed.
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
  cloneParagraphWithoutPropertySource,
  cloneParagraphWithPropertySource,
  copyParagraphIndentationProjection,
  copyParagraphPropertyCapture,
  copyParagraphPropertySource,
  createProseParagraphWithPropertySource,
  linkParagraphPropertySourceCandidate,
  paragraphPropertySourceState,
  restoreParagraphIndentationProjection,
  sealParagraphFacts,
  transferProseParagraphPropertySource,
} from "./paragraphPropertySource";

setDefaultTimeout(propertyTestTimeout(60_000));

const POOL = 5;

const stepArbitrary = fc.record({
  kind: fc.constantFrom(
    "seal",
    "source",
    "copySource",
    "copyCapture",
    "copyProjection",
    "restoreProjection",
    "assignProjection",
    "spread",
    "cloneWithout",
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

const run = (
  step: Step,
  pool: Paragraph[],
  nodes: PMNode[],
  sealed: Set<Paragraph>,
): void => {
  const target = pool[step.target] ?? pool[0];
  const source = pool[step.source] ?? pool[0];
  if (target === undefined || source === undefined) return;
  switch (step.kind) {
    case "seal":
      sealParagraphFacts(target);
      sealed.add(target);
      return;
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
    case "cloneWithout":
      pool[step.target] = cloneParagraphWithoutPropertySource(source, {});
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

const REFUSAL = /^A paragraph fact \(.+\) changed after a reader kept it\.$/;

describe("sealed paragraph facts", () => {
  test("a writer never changes the facts of a sealed paragraph", () => {
    let refusals = 0;
    let unsealedWrites = 0;
    assertProperty(
      fc.property(
        fc.array(
          fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
          { minLength: POOL, maxLength: POOL },
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
          const sealed = new Set<Paragraph>();
          for (const step of steps) {
            const before = [...sealed].map((paragraph) => ({
              paragraph,
              formatting: paragraph.formatting,
              state: paragraphPropertySourceState(paragraph),
            }));
            const unsealed = pool
              .filter((paragraph) => !sealed.has(paragraph))
              .map((paragraph) => paragraphPropertySourceState(paragraph));
            try {
              run(step, pool, nodes, sealed);
            } catch (error) {
              const message = error instanceof Error ? error.message : "";
              if (!REFUSAL.test(message)) throw error;
              refusals += 1;
            }
            for (const { paragraph, formatting, state } of before) {
              expect(paragraph.formatting).toBe(formatting);
              const now = paragraphPropertySourceState(paragraph);
              expect(now.length).toBe(state.length);
              for (const [index, fact] of now.entries())
                expect(fact).toBe(state[index]);
            }
            const after = pool
              .filter((paragraph) => !sealed.has(paragraph))
              .map((paragraph) => paragraphPropertySourceState(paragraph));
            if (
              after.length === unsealed.length &&
              after.some((state, index) =>
                state.some((fact, at) => fact !== unsealed[index]?.[at]),
              )
            )
              unsealedWrites += 1;
          }
        },
      ),
      { id: "sealed paragraph facts never change", numRuns: 500 },
    );
    // The writers ran against both sealed and open paragraphs.
    expect(refusals).toBeGreaterThan(0);
    expect(unsealedWrites).toBeGreaterThan(0);
  });
});
