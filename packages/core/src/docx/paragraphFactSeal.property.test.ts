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
 * and a refusal must name the fact it would have changed. Carrying one
 * paragraph's facts onto another never refuses: it gives the target the
 * facts the copy makes, on a new paragraph when the target is kept.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { checkCoverageSample } from "../../../../test/property-coverage";
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
  carryParagraphPropertySource,
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
    "carry",
    "carryRestore",
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

/**
 * The facts `carryParagraphPropertySource` gives `target`, in the order
 * `paragraphPropertySourceState` reads them: those `source` holds, the
 * target's own where the source holds none, and the target's projection
 * unless it holds none or history restores the source's.
 */
const carriedFacts = (
  target: Paragraph,
  source: Paragraph,
  indentation: "derive" | "restore",
): readonly unknown[] => {
  const [
    token,
    candidate,
    transferId,
    kind,
    capture,
    projection,
    assigned,
    payload,
    ...byFormatting
  ] = paragraphPropertySourceState(target);
  const [
    sourceToken,
    sourceCandidate,
    sourceTransferId,
    sourceKind,
    sourceCapture,
    sourceProjection,
  ] = paragraphPropertySourceState(source);
  const keepsProjection = assigned === true || projection !== undefined;
  return [
    sourceToken ?? token,
    sourceCandidate ?? candidate,
    sourceTransferId ?? transferId,
    kind ??
      (sourceToken !== undefined || sourceKind === "durable"
        ? "durable"
        : "capture-only"),
    sourceCapture ?? capture,
    indentation === "derive" && keepsProjection ? projection : sourceProjection,
    indentation === "restore" ? true : assigned,
    payload,
    ...byFormatting,
  ];
};

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
    case "carry":
    case "carryRestore": {
      const indentation = step.kind === "carry" ? "derive" : "restore";
      const expected = carriedFacts(target, source, indentation);
      const carried = carryParagraphPropertySource({
        target,
        source,
        indentation,
      });
      const facts = paragraphPropertySourceState(carried);
      expect(facts.length).toBe(expected.length);
      for (const [index, fact] of facts.entries())
        expect(fact).toBe(expected[index]);
      expect(carried.formatting).toBe(target.formatting);
      // Only a kept paragraph is replaced, and only by one that differs.
      if (carried !== target) {
        expect(sealed.has(target)).toBe(true);
        pool[step.target] = carried;
      }
      return;
    }
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
    type Counts = {
      refusals: number;
      unsealedWrites: number;
      replaced: number;
    };
    const law = (counts: Counts) =>
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
            const held = pool[step.target];
            try {
              run(step, pool, nodes, sealed);
            } catch (error) {
              const message = error instanceof Error ? error.message : "";
              // Carrying never refuses: it replaces a kept paragraph instead.
              if (
                !REFUSAL.test(message) ||
                step.kind === "carry" ||
                step.kind === "carryRestore"
              )
                throw error;
              counts.refusals += 1;
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
              counts.unsealedWrites += 1;
            if (
              (step.kind === "carry" || step.kind === "carryRestore") &&
              pool[step.target] !== held
            )
              counts.replaced += 1;
          }
        },
      );
    assertProperty(law({ refusals: 0, unsealedWrites: 0, replaced: 0 }), {
      id: "sealed paragraph facts never change",
      numRuns: 500,
    });
    // On a fixed sample, the writers ran against both sealed and open
    // paragraphs, and carrying replaced kept ones. Each case runs up to 30
    // steps, so fewer cases than the property's reach every outcome.
    const RUNS = 200;
    const counts: Counts = { refusals: 0, unsealedWrites: 0, replaced: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.refusals).toBeGreaterThan(0);
    expect(counts.unsealedWrites).toBeGreaterThan(0);
    expect(counts.replaced).toBeGreaterThan(0);
  });

  test("each writer leaves a paragraph sealed just before it as it was", () => {
    assertProperty(
      fc.property(
        fc.array(
          fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
          { minLength: POOL, maxLength: POOL },
        ),
        fc.array(stepArbitrary, { maxLength: 15 }),
        stepArbitrary,
        (indents, history, last) => {
          const pool = indents.map((indent, index) =>
            cloneParagraphWithPropertySource(paragraphAt(index, indent), {}),
          );
          const nodes: PMNode[] = [];
          const sealed = new Set<Paragraph>();
          const attempt = (step: Step): void => {
            try {
              run(step, pool, nodes, sealed);
            } catch (error) {
              const message = error instanceof Error ? error.message : "";
              if (
                !REFUSAL.test(message) ||
                step.kind === "carry" ||
                step.kind === "carryRestore"
              )
                throw error;
            }
          };
          for (const step of history) attempt(step);
          const target = pool[last.target];
          if (target === undefined) return;
          sealParagraphFacts(target);
          sealed.add(target);
          const formatting = target.formatting;
          const state = paragraphPropertySourceState(target);
          attempt(last);
          expect(target.formatting).toBe(formatting);
          const now = paragraphPropertySourceState(target);
          expect(now.length).toBe(state.length);
          for (const [index, fact] of now.entries())
            expect(fact).toBe(state[index]);
        },
      ),
      { id: "a writer leaves a just-sealed paragraph", numRuns: 1000 },
    );
  });

  test("every writer refuses a kept paragraph whose facts it would change", () => {
    const withProjection = (): Paragraph => {
      const paragraph = paragraphAt(1, undefined);
      // Inherited indentation the paragraph does not author keeps a projection.
      assignParagraphIndentationProjection({
        paragraph,
        authored: undefined,
        inherited: { indentLeft: 360 },
      });
      return paragraph;
    };
    const withSource = (): Paragraph => {
      const paragraph = paragraphAt(2, undefined);
      assignParagraphPropertySource(paragraph, "<w:pPr/>");
      return paragraph;
    };
    const writers: Record<string, (target: Paragraph) => void> = {
      source: (target) => assignParagraphPropertySource(target, "<w:pPr/>"),
      copySource: (target) => copyParagraphPropertySource(target, withSource()),
      copyCapture: (target) =>
        copyParagraphPropertyCapture(target, withSource()),
      copyProjection: (target) =>
        copyParagraphIndentationProjection({
          target,
          source: withProjection(),
        }),
      restoreProjection: (target) =>
        restoreParagraphIndentationProjection({
          target,
          source: withProjection(),
        }),
      assignProjection: (target) =>
        assignParagraphIndentationProjection({
          paragraph: target,
          authored: undefined,
          inherited: { indentLeft: 720 },
        }),
      candidate: (target) =>
        linkParagraphPropertySourceCandidate(
          target,
          createProseParagraphWithPropertySource(
            schema.nodes["paragraph"],
            withSource(),
          ),
        ),
      contract: (target) =>
        assignDocumentParagraphPropertySourceContract(
          { package: { document: { content: [target] } } },
          DIGEST,
        ),
    };
    for (const [kind, write] of Object.entries(writers)) {
      const target = paragraphAt(0, 1);
      sealParagraphFacts(target);
      const formatting = target.formatting;
      const state = paragraphPropertySourceState(target);
      expect(() => write(target), kind).toThrow(REFUSAL);
      expect(target.formatting).toBe(formatting);
      expect(paragraphPropertySourceState(target)).toEqual(state);
    }
  });
});
