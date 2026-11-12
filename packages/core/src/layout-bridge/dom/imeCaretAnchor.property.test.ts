/**
 * The IME anchor measures the hidden caret with the host's translation in
 * place and takes that translation off, instead of clearing the transform to
 * measure. Over any sequence of syncs and resets, the offset it leaves must
 * equal the one measured from the untransformed position, it must write only
 * when that offset changes, and once the host has a translation it must keep
 * one: removing and restoring the transform rebuilds paint properties for the
 * whole hidden document.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { resetImeCaretAnchor, syncImeCaretAnchor } from "./imeCaretAnchor";

setDefaultTimeout(propertyTestTimeout(30_000));

type Point = { left: number; top: number };

/** Offsets the way a browser applies the host's inline transform. */
const appliedOffset = (transform: string): { dx: number; dy: number } => {
  if (transform === "") return { dx: 0, dy: 0 };
  const match = /^translate3d\((-?[\d.]+)px, (-?[\d.]+)px, 0(?:px)?\)$/u.exec(
    transform,
  );
  if (!match) throw new Error(`unexpected transform ${transform}`);
  return { dx: Number(match[1]), dy: Number(match[2]) };
};

/** A host whose transform writes are counted. */
const countingHost = () => {
  let transform = "";
  let writes = 0;
  const host = {
    style: {
      get transform() {
        return transform;
      },
      set transform(value: string) {
        transform = value;
        writes += 1;
      },
    },
  };
  return { host, writes: () => writes };
};

// Layout positions are multiples of 1/64 px, so the model's sums and
// differences are exact and rounding ties cannot drift.
const coordinate = fc
  .integer({ min: -20_000 * 64, max: 20_000 * 64 })
  .map((units) => units / 64);
const point = fc.record({ left: coordinate, top: coordinate });

type Step =
  | {
      kind: "sync";
      hidden: Point;
      visible: Point | null;
      composing: boolean;
      focused: boolean;
      empty: boolean;
    }
  | { kind: "reset" };

const step: fc.Arbitrary<Step> = fc.oneof(
  {
    weight: 6,
    arbitrary: fc.record({
      kind: fc.constant("sync" as const),
      hidden: point,
      visible: fc.option(point, { freq: 4 }),
      composing: fc.oneof(
        { weight: 6, arbitrary: fc.constant(false) },
        { weight: 1, arbitrary: fc.constant(true) },
      ),
      focused: fc.oneof(
        { weight: 6, arbitrary: fc.constant(true) },
        { weight: 1, arbitrary: fc.constant(false) },
      ),
      empty: fc.oneof(
        { weight: 6, arbitrary: fc.constant(true) },
        { weight: 1, arbitrary: fc.constant(false) },
      ),
    }),
  },
  { weight: 1, arbitrary: fc.constant({ kind: "reset" as const }) },
);

// A caret that sits still across several syncs, the common case while typing.
const repeated = fc
  .tuple(step, fc.integer({ min: 1, max: 3 }))
  .map(([one, times]) => Array.from({ length: times }, () => one));
const steps = fc.array(repeated, { maxLength: 12 }).map((runs) => runs.flat());

describe("syncImeCaretAnchor", () => {
  test("leaves the offset measured from the untransformed caret, writing only on change and never removing the translation", () => {
    fc.assert(
      fc.property(steps, (sequence) => {
        const { host, writes } = countingHost();
        // The reference: the offset from the caret's untransformed position.
        let expected = { dx: 0, dy: 0 };
        let translated = false;
        for (const current of sequence) {
          const before = writes();
          const previous = expected;
          if (current.kind === "reset") {
            resetImeCaretAnchor(host);
            expected = { dx: 0, dy: 0 };
          } else {
            const { hidden, visible, composing, focused, empty } = current;
            syncImeCaretAnchor({
              hiddenHost: host,
              editorView: {
                composing,
                hasFocus: () => focused,
                state: { selection: { empty, head: 1 } },
                // The hidden caret moves with the host's translation.
                coordsAtPos: () => {
                  const offset = appliedOffset(host.style.transform);
                  return {
                    left: hidden.left + offset.dx,
                    top: hidden.top + offset.dy,
                  };
                },
              },
              visibleCaret: visible,
            });
            // A live composition keeps its anchor untouched.
            if (composing) {
              expect(writes()).toBe(before);
              continue;
            }
            expected =
              visible && focused && empty
                ? {
                    // `+ 0` folds a rounded -0 into 0, as a written px value does.
                    dx: Math.round(visible.left - hidden.left) + 0,
                    dy: Math.round(visible.top - hidden.top) + 0,
                  }
                : { dx: 0, dy: 0 };
          }
          expect(appliedOffset(host.style.transform)).toEqual(expected);
          if (expected.dx === previous.dx && expected.dy === previous.dy)
            expect(writes()).toBe(before);
          if (translated) expect(host.style.transform).not.toBe("");
          translated ||= host.style.transform !== "";
        }
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });
});
