/**
 * CPU of a markup-view switch on a deterministic 500-page review document.
 * Pass --compare to interleave the shipped projector and the reference projector.
 * Layout uses deterministic text measurement and stops before browser painting.
 */
import { loadavg } from "node:os";
import { EditorState } from "prosemirror-state";

import { normalizeForOps, validateOpsDocument } from "@stll/docx-core/ops";

import { remapMarkupViewBlocks } from "../src/layout-bridge/convert/markupViewFlow";
import { toFlowBlocks } from "../src/layout-bridge/convert/toFlowBlocks";
import { layoutDocument } from "../src/layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../src/layout-engine/measure/__tests__/fakeTextMeasure";
import { clearAllCaches } from "../src/layout-engine/measure/cache";
import { measureBlocks } from "../src/layout-engine/measure/measureBlocks";
import type { DisplayMode } from "../src/managers/EditorModeManager";
import {
  ProjectionBlockCache,
  toProseDoc,
} from "../src/prosemirror/conversion/toProseDoc";
import {
  type MarkupViewProjection,
  projectMarkupView,
} from "../src/prosemirror/markupViewProjection";
import type { Paragraph, Run } from "../src/types/document";
import { createEmptyDocument } from "../src/utils/createDocument";

const PAGE_COUNT = 500;
const PARAGRAPHS_PER_PAGE = 8;
const RUNS_PER_PARAGRAPH = 5;
const WARMUPS = 2;
const SAMPLES = 9;
const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const VIEWS = [
  "no-markup",
  "original",
  "simple-markup",
] as const satisfies readonly DisplayMode[];

const line = (page: number, paragraph: number, index: number): Run => ({
  type: "run",
  content: [
    {
      type: "text",
      text: `Page ${page + 1}, paragraph ${paragraph + 1}, run ${index + 1}: the parties confirm the terms and conditions. `,
    },
  ],
});

const document = createEmptyDocument();
document.package.document.content = Array.from(
  { length: PAGE_COUNT * PARAGRAPHS_PER_PAGE },
  (_, index): Paragraph => {
    const page = Math.floor(index / PARAGRAPHS_PER_PAGE);
    const paragraph = index % PARAGRAPHS_PER_PAGE;
    const startsPage = paragraph === 0;
    const paraId = (index + 1).toString(16).padStart(8, "0").toUpperCase();
    const runs = Array.from({ length: RUNS_PER_PARAGRAPH }, (_unused, run) =>
      line(page, paragraph, run),
    );
    const author = "Reviewer";
    const format =
      page > 0 && startsPage ? { formatting: { pageBreakBefore: true } } : {};
    if (!startsPage) return { type: "paragraph", paraId, content: runs };
    const [first, removed, inserted, ...rest] = runs;
    if (first === undefined || removed === undefined || inserted === undefined)
      throw new Error(
        "Review benchmark requires at least three runs per paragraph",
      );
    return {
      type: "paragraph",
      paraId,
      ...format,
      content: [
        first,
        {
          type: "deletion",
          info: { author, id: page * 2 + 1 },
          content: [removed],
        },
        {
          type: "insertion",
          info: { author, id: page * 2 + 2 },
          content: [inserted],
        },
        ...rest,
      ],
    };
  },
);

const normalizedDocument = normalizeForOps(document);
const doc = toProseDoc(normalizedDocument);
const current = { implementation: "current", project: projectMarkupView };
const compare = process.argv.includes("--compare");
const implementations = compare
  ? [
      current,
      {
        implementation: "reference",
        project: (
          await import("../src/__tests__/reference-editor/markupViewProjection")
        ).projectReferenceMarkupView,
      },
    ]
  : [current];

const measure = <T>(operation: () => T) => {
  const started = process.cpuUsage();
  const value = operation();
  const cpu = process.cpuUsage(started);
  return { value, cpuMs: (cpu.user + cpu.system) / 1000 };
};

const cpuSummary = (samples: number[]) => {
  const sorted = samples.toSorted((left, right) => left - right);
  return {
    median: sorted.at(Math.floor(sorted.length / 2)),
    samples,
  };
};

const sourceFixtures = [
  {
    source: "model",
    freshProjection: () => {
      const source = validateOpsDocument({
        package: {
          ...normalizedDocument.package,
          document: { ...normalizedDocument.package.document },
        },
      }).unwrap();
      return toProseDoc(source, { blockCache: new ProjectionBlockCache() });
    },
  },
  {
    source: "detached",
    freshProjection: () => doc.type.create(doc.attrs, doc.content, doc.marks),
  },
];

const layoutView = (projection: MarkupViewProjection) => {
  if (projection.type !== "resolved")
    throw new Error("Expected a resolved view for the tracked fixture");
  const blocks = remapMarkupViewBlocks(
    toFlowBlocks(projection.doc),
    projection,
  );
  return layoutDocument(blocks, measureBlocks(blocks, CONTENT_WIDTH), {
    pageSize: PAGE_SIZE,
    margins: MARGINS,
  });
};

const sampleSwitch = ({
  view,
  project,
  freshProjection,
}: {
  view: (typeof VIEWS)[number];
  project: typeof projectMarkupView;
  freshProjection: (typeof sourceFixtures)[number]["freshProjection"];
}) => {
  // Loading the document and creating its editor state happen before a switch.
  const state = EditorState.create({ doc: freshProjection() });
  project(state, "all-markup");
  const projection = measure(() => project(state, view));
  // The two engines share global layout caches. Give each arm the same
  // starting cache so the engine sampled second cannot reuse the first arm.
  clearAllCaches();
  const layout = measure(() => layoutView(projection.value));
  if (layout.value.pages.length !== PAGE_COUNT)
    throw new Error(
      `Expected ${PAGE_COUNT} pages in ${view}, got ${layout.value.pages.length}`,
    );
  const warmLayout = measure(() => {
    project(state, "all-markup");
    return layoutView(project(state, view));
  });
  const cached = measure(() => {
    for (let switchIndex = 0; switchIndex < 1000; switchIndex += 1) {
      project(state, "all-markup");
      project(state, view);
    }
  });
  return {
    coldProjectionCpuMs: projection.cpuMs,
    coldProjectionAndLayoutCpuMs: projection.cpuMs + layout.cpuMs,
    warmProjectionAndLayoutCpuMs: warmLayout.cpuMs,
    cachedProjectionCpuMs: cached.cpuMs / 1000,
  };
};

const loadBefore = loadavg().at(0);
const results: object[] = [];
withFakeTextMeasure(
  () => {
    for (const { source, freshProjection } of sourceFixtures) {
      for (const view of VIEWS) {
        const groups = implementations.map(({ implementation, project }) => ({
          implementation,
          project,
          measurements: Array<ReturnType<typeof sampleSwitch>>(),
        }));
        for (let sample = 0; sample < WARMUPS + SAMPLES; sample += 1) {
          // Alternate which implementation runs first to balance timing bias
          // from shared measurement caches and garbage collection.
          const sampleOrder = sample % 2 === 0 ? groups : groups.toReversed();
          for (const group of sampleOrder) {
            const measurement = sampleSwitch({
              view,
              project: group.project,
              freshProjection,
            });
            if (sample >= WARMUPS) group.measurements.push(measurement);
          }
        }
        for (const group of groups) {
          results.push({
            source,
            view,
            implementation: group.implementation,
            coldProjectionCpuMs: cpuSummary(
              group.measurements.map((sample) => sample.coldProjectionCpuMs),
            ),
            coldProjectionAndLayoutCpuMs: cpuSummary(
              group.measurements.map(
                (sample) => sample.coldProjectionAndLayoutCpuMs,
              ),
            ),
            warmProjectionAndLayoutCpuMs: cpuSummary(
              group.measurements.map(
                (sample) => sample.warmProjectionAndLayoutCpuMs,
              ),
            ),
            cachedProjectionCpuMs: cpuSummary(
              group.measurements.map((sample) => sample.cachedProjectionCpuMs),
            ),
          });
        }
      }
    }
  },
  { charWidth: fixedCharWidth(5) },
);
console.log(
  JSON.stringify(
    {
      pages: PAGE_COUNT,
      sourceCharacters: doc.textContent.length,
      trackedRevisions: PAGE_COUNT * 2,
      loadAverage: { before: loadBefore, after: loadavg().at(0) },
      results,
    },
    null,
    2,
  ),
);
