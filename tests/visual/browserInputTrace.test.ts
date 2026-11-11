import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";

import {
  documentShape,
  shapeArrayBuffer,
} from "../../packages/core/src/__tests__/documentShapes";
import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { toProseDoc } from "../../packages/core/src/prosemirror/conversion/toProseDoc";
import {
  propertyConfig,
  propertyTestTimeout,
} from "../../test/property-testing";
import {
  type BrowserDragTarget,
  crossParagraphDragRange,
} from "./browserDragTarget";
import {
  BROWSER_INPUT_ACTION_DISPOSITIONS,
  BROWSER_PARAGRAPH_DRAG_SHAPES,
  BROWSER_SHAPE_TARGETS,
  BROWSER_TRACE_FIXED_SEEDS,
  browserImeActionArbitrary,
  browserInputTraceArbitrary,
  browserSuggestionActionKinds,
  parseBrowserInputTraceConfig,
} from "./browserInputTrace";

setDefaultTimeout(propertyTestTimeout(5_000));

test("browser traces replay deterministically and exercise the declared shapes and input kinds", () => {
  const arbitrary = browserInputTraceArbitrary(12);
  const config = { seed: 197, numRuns: 800 };
  const traces = fc.sample(arbitrary, config);
  expect(fc.sample(arbitrary, config)).toEqual(traces);
  expect(new Set(traces.map(({ shape }) => shape))).toEqual(
    new Set(Object.keys(BROWSER_SHAPE_TARGETS)),
  );
  expect(new Set(browserSuggestionActionKinds)).toEqual(
    new Set(SUGGESTION_INPUT_KINDS),
  );
  const kinds = new Set(
    traces.flatMap(({ actions }) => actions.map(({ kind }) => kind)),
  );
  for (const kind of [
    ...SUGGESTION_INPUT_KINDS,
    "undo",
    "redo",
    "historyBurst",
    "selectionDrag",
    "selectAll",
  ]) {
    expect(kinds.has(kind)).toBe(true);
  }
  expect([...kinds].sort()).toEqual(
    Object.keys(BROWSER_INPUT_ACTION_DISPOSITIONS)
      .filter((kind) => kind !== "keys")
      .sort(),
  );
  for (const action of traces.flatMap(({ actions }) => actions)) {
    if (
      action.kind === "undo" ||
      action.kind === "redo" ||
      action.kind === "historyBurst"
    )
      expect(BROWSER_INPUT_ACTION_DISPOSITIONS[action.kind]).toBe("history");
    else if (action.kind === "selectionDrag" || action.kind === "selectAll")
      expect(BROWSER_INPUT_ACTION_DISPOSITIONS[action.kind]).toBe("selection");
    else expect(BROWSER_INPUT_ACTION_DISPOSITIONS[action.kind]).toBe("edit");
  }
  const targetFeatures = {
    table: "table",
    list: "list-decimal",
    note: "footnote",
    field: "field",
    inlineObject: "image",
  } as const satisfies Record<Exclude<BrowserDragTarget, "paragraphs">, string>;
  const shapesWithParagraphDrags = new Set<string>();
  for (const { shape, actions } of traces) {
    const fixture = documentShape(shape);
    for (const action of actions) {
      if (action.kind === "selectionDrag" && action.target === "paragraphs") {
        expect(BROWSER_PARAGRAPH_DRAG_SHAPES.has(shape)).toBe(true);
        shapesWithParagraphDrags.add(shape);
      } else if (action.kind === "selectionDrag") {
        expect(action.target).toBe(BROWSER_SHAPE_TARGETS[shape]);
        if (action.target === "list") {
          expect(
            fixture.features.some(
              (feature) =>
                feature === "list-decimal" || feature === "list-bullet",
            ),
          ).toBe(true);
        } else {
          expect(fixture.features).toContain(targetFeatures[action.target]);
        }
      }
      if (action.kind === "dragCellDelete")
        expect(fixture.features).toContain("table");
    }
  }
  expect(shapesWithParagraphDrags).toEqual(
    new Set(BROWSER_PARAGRAPH_DRAG_SHAPES),
  );
});

test("a paragraphs drag is offered exactly for shapes with two adjacent text paragraphs", async () => {
  for (const shape of Object.keys(BROWSER_SHAPE_TARGETS)) {
    // oxlint-disable-next-line no-await-in-loop -- one fixture at a time
    const doc = toProseDoc(
      // oxlint-disable-next-line no-await-in-loop -- one fixture at a time
      await parseDocx(await shapeArrayBuffer(shape), {
        preloadFonts: false,
        detectVariables: false,
      }),
    );
    const range = crossParagraphDragRange(doc);
    expect([shape, range !== null]).toEqual([
      shape,
      BROWSER_PARAGRAPH_DRAG_SHAPES.has(shape as never),
    ]);
    if (range === null) continue;
    const $from = doc.resolve(range.from);
    const $to = doc.resolve(range.to);
    expect($from.parent.isTextblock && $to.parent.isTextblock).toBe(true);
    expect($from.sameParent($to)).toBe(false);
    expect($to.index(0)).toBe($from.index(0) + 1);
  }
});

test("a paragraphs drag passes over struck text, as the applied input reads", async () => {
  const doc = toProseDoc(
    await parseDocx(await shapeArrayBuffer("single-bullet-list"), {
      preloadFonts: false,
      detectVariables: false,
    }),
  );
  const range = crossParagraphDragRange(doc);
  if (range === null) throw new TypeError("The list shape has no drag.");
  const $first = doc.resolve(range.from);
  const deletion = doc.type.schema.marks["deletion"]?.create({ revisionId: 1 });
  if (deletion === undefined) throw new TypeError("No deletion mark.");
  // The first paragraph struck whole reads as no text, so the drag starts in
  // the next pair, as it would once the deletion is applied.
  const struckChildren: PMNode[] = [];
  $first.parent.forEach((child) => {
    struckChildren.push(child.mark(deletion.addToSet(child.marks)));
  });
  const struck = doc.replace(
    $first.start(),
    $first.end(),
    new Slice(Fragment.fromArray(struckChildren), 0, 0),
  );
  const shifted = crossParagraphDragRange(struck);
  if (shifted === null) throw new TypeError("The struck shape has no drag.");
  const $from = struck.resolve(shifted.from);
  const $to = struck.resolve(shifted.to);
  expect($from.index(0)).toBe($first.index(0) + 1);
  expect($to.index(0)).toBe($from.index(0) + 1);
  expect(shifted.from).toBe($from.start() + 1);
  expect(shifted.to).toBe($to.start() + 2);
});

test("IME generator covers cancellation, repeated updates and Unicode with valid shrinking", () => {
  const actions = fc.sample(browserImeActionArbitrary, {
    seed: 431,
    numRuns: 300,
  });
  expect(new Set(actions.map(({ completion }) => completion))).toEqual(
    new Set(["commit", "cancel"]),
  );
  expect(actions.some(({ updates }) => updates.length > 1)).toBe(true);
  for (const text of ["東京", "e\u0301", "👩🏽‍⚖️", "مرحبا"]) {
    expect(
      actions.some(({ updates }) =>
        updates.some((update) => update.includes(text)),
      ),
    ).toBe(true);
  }
  // An intentionally failing oracle proves lifecycle payloads shrink, rather than
  // treating each sampled multi-update action as an opaque constant.
  const failure = fc.check(
    fc.property(browserImeActionArbitrary, () => false),
    propertyConfig({ seed: 431 }),
  );
  expect(failure.failed).toBe(true);
  const minimal = failure.counterexample?.at(0);
  expect(minimal?.updates).toHaveLength(1);
  expect(minimal?.updates.at(0)).toBe("alpha");
  expect(minimal?.completion).toBe("commit");
});

test("browser trace sequences shrink to one action", () => {
  const failure = fc.check(
    fc.property(browserInputTraceArbitrary(12), () => false),
    propertyConfig({ seed: 197 }),
  );
  expect(failure.failed).toBe(true);
  expect(failure.counterexample?.at(0)?.actions).toHaveLength(1);
});

test("browser lanes retain pinned seeds and reject malformed overrides", () => {
  expect(parseBrowserInputTraceConfig({}, "nightly")).toEqual({
    seeds: BROWSER_TRACE_FIXED_SEEDS.nightly,
    defaultEditorSeeds: [],
    runs: 20,
  });
  expect(parseBrowserInputTraceConfig({}, "pullRequest")).toEqual({
    seeds: BROWSER_TRACE_FIXED_SEEDS.pullRequest,
    defaultEditorSeeds: [],
    runs: 2,
  });
  expect(
    parseBrowserInputTraceConfig(
      { FOLIO_FUZZ_SEEDS: "11,-29", FOLIO_FUZZ_RUNS: "7" },
      "nightly",
    ),
  ).toEqual({ seeds: [11, -29], defaultEditorSeeds: [], runs: 7 });
  expect(
    parseBrowserInputTraceConfig(
      { FOLIO_FUZZ_DEFAULT_EDITOR_SEEDS: "29,83" },
      "nightly",
    ).defaultEditorSeeds,
  ).toEqual([29, 83]);
  for (const seeds of ["", " ", "11,", ",29", "11,,29", "NaN", "1.5"]) {
    expect(() =>
      parseBrowserInputTraceConfig({ FOLIO_FUZZ_SEEDS: seeds }, "nightly"),
    ).toThrow();
    expect(() =>
      parseBrowserInputTraceConfig(
        { FOLIO_FUZZ_DEFAULT_EDITOR_SEEDS: seeds },
        "nightly",
      ),
    ).toThrow();
  }
  for (const runs of ["", "0", "-1", "1.5", "NaN"]) {
    expect(() =>
      parseBrowserInputTraceConfig({ FOLIO_FUZZ_RUNS: runs }, "nightly"),
    ).toThrow();
  }
});
