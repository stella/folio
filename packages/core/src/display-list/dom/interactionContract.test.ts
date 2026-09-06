/**
 * Every hook the interaction layer reads has a source in the paint IR.
 *
 * The surface that edits a page finds things in the DOM by class name and by
 * data attribute. Those names were the painter's, and every reader learned
 * them; a renderer that did not emit one of them silently stopped answering a
 * question the editor asks — a click that lands nowhere, a header that cannot
 * be entered, a comment that cannot be found.
 *
 * So the contract is enumerated here rather than remembered: the test reads the
 * interaction layer's own source, extracts every hook it looks for, and fails
 * unless each one is accounted for. Adding a selector to a reader without
 * giving the producer something to emit it from fails this test, which is the
 * only way "the DOM contract" stays a contract rather than a habit.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";

/** The modules that hit-test the painted page. */
const READERS = [
  "../../layout-bridge/dom/clickToPositionDom.ts",
  "../../layout-bridge/dom/findBodyPmSpans.ts",
  "../../layout-bridge/dom/findHfPmSpans.ts",
  "../../layout-bridge/dom/noteStoryDom.ts",
  "../../layout-bridge/dom/imeCaretAnchor.ts",
] as const;

/**
 * What emits each hook.
 *
 * `region` names the region kind whose element carries it, `run` the glyph-run
 * element, and `container` the page shell the editor owns rather than the
 * producer. A hook with no source is a question the display list cannot answer.
 */
const SOURCES = {
  "layout-page": "container",
  "layout-page-content": "region: pageContent",
  "layout-page-header": "region: headerSlot",
  "layout-page-footer": "region: footerSlot",
  "layout-paragraph": "region: paragraph",
  "layout-line": "region: line",
  "layout-empty-run": "region: emptyRun",
  "layout-table": "region: table",
  "layout-table-cell": "region: tableCell",
  "layout-run-text": "run",
  "layout-run-tab": "region: tab",
  "layout-document": "container",
  "pm-start": "run and region: model.pmRange.start",
  "pm-end": "run and region: model.pmRange.end",
  "page-number": "container",
  rid: "region: headerSlot/footerSlot, from model.story.rId",
  "hf-rid": "region: headerSlot/footerSlot, from model.story.rId",
  "hf-r-id": "region: headerSlot/footerSlot, from model.story.rId",
  "hf-slot-kind": "region: headerSlot/footerSlot, from model.story.kind",
  "hf-kind": "region: headerSlot/footerSlot, from model.story.kind",
  "note-kind": "region: note, from model.story.kind",
  "note-id": "region: note, from model.story.id",
  "comment-id": "region: line, from model.commentIds",
  "block-id": "region, from model.blockId",
  "row-index": "region: tableRow, from model.rowIndex",
  "column-index": "region: tableCell, from model.columnIndex",
  story: "run and region, from model.story.kind",
  "advance-sum": "run: the producer's own advance total",
  "collapsed-leading-spaces": "run: collapsedEdge.side",
  "collapsed-trailing-spaces": "run: collapsedEdge.side",
  "collapsed-space-advance": "run: collapsedEdge.spaceAdvancePx",
} as const;

const sourceOf = (reader: string): string => readFileSync(new URL(reader, import.meta.url), "utf8");

/** Class names a reader looks for, however it spells the lookup. */
const classHooks = (source: string): readonly string[] => [
  ...new Set(
    [
      ...source.matchAll(/classList\.contains\("([a-z-]+)"\)/gu),
      ...source.matchAll(/["'`.]\.?(layout-[a-z-]+)/gu),
    ].flatMap(([, name]) => (name === undefined ? [] : [name])),
  ),
];

/** Data attributes a reader looks for, as their attribute spelling. */
const dataHooks = (source: string): readonly string[] => [
  ...new Set(
    [
      ...source.matchAll(/dataset\["([a-zA-Z]+)"\]/gu),
      ...source.matchAll(/\[data-([a-z-]+)[\]=]/gu),
    ].flatMap(([, name]) =>
      name === undefined ? [] : [name.replaceAll(/[A-Z]/gu, (upper) => `-${upper.toLowerCase()}`)],
    ),
  ),
];

describe("the interaction layer reads nothing the display list cannot emit", () => {
  for (const reader of READERS) {
    test(`${reader.split("/").at(-1) ?? reader} reads only accounted-for hooks`, () => {
      const source = sourceOf(reader);
      const hooks = [...classHooks(source), ...dataHooks(source)];
      const unaccounted = hooks.filter((hook) => !(hook in SOURCES));

      expect({ reader, unaccounted }).toEqual({ reader, unaccounted: [] });
    });
  }

  test("every accounted hook the display list emits is emitted by name", () => {
    const backend = sourceOf("./renderDisplayListToDom.ts");
    const missing = Object.entries(SOURCES).flatMap(([hook, source]) => {
      if (source === "container" || source.startsWith("not emitted")) {
        return [];
      }
      // A class is emitted verbatim; a data attribute is written through
      // `dataset`, which spells it in camel case.
      const camel = hook.replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());
      return backend.includes(`"${hook}"`) || backend.includes(`"${camel}"`) ? [] : [hook];
    });

    expect(missing).toEqual([]);
  });

  test("no hook is left without a source", () => {
    // Stated as an empty list rather than left implicit: a hook the display
    // list cannot answer is a question the editor asks and gets nothing back
    // for, and adding one has to be a deliberate edit here.
    const gaps = Object.entries(SOURCES)
      .filter(([, source]) => source.startsWith("not emitted"))
      .map(([hook]) => hook);

    expect(gaps).toEqual([]);
  });
});
