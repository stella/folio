import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Paragraph,
  Run,
  RunContent,
  Table,
  TableCell,
  TableRow,
} from "../types/document";
import { buildPageBreakRunSourceDescendantIndex } from "./pageBreakRunSourceDescendantIndex";

type SourceReadObservation = {
  nodeReads: Map<string, number>;
  runReads: Map<string, number>;
};

type NestedSource = {
  content: BlockContent[];
  ownerContents: readonly BlockContent[][];
  expectedNodeReads: number;
  expectedRunReads: number;
  observation: SourceReadObservation;
};

type ObserveTraversalReadOptions<T extends object> = {
  node: T;
  property: keyof T;
  label: string;
  observation: SourceReadObservation;
  kind?: "node" | "run";
};

const observeTraversalRead = <T extends object>({
  node,
  property: observedProperty,
  label,
  observation,
  kind = "node",
}: ObserveTraversalReadOptions<T>): T =>
  new Proxy(node, {
    get: (target, property, receiver) => {
      if (property === observedProperty) {
        observation.nodeReads.set(label, (observation.nodeReads.get(label) ?? 0) + 1);
        if (kind === "run") {
          observation.runReads.set(label, (observation.runReads.get(label) ?? 0) + 1);
        }
      }
      return Reflect.get(target, property, receiver);
    },
  });

const pageBreakRun = (): Run => ({
  type: "run",
  content: [{ type: "break", breakType: "page" }],
});

const pageBreakParagraph = (): Paragraph => ({ type: "paragraph", content: [pageBreakRun()] });

const textBoxRun = (content: BlockContent[]): Run => ({
  type: "run",
  content: [
    {
      type: "shape",
      shape: {
        type: "shape",
        shapeType: "rect",
        size: { width: 914_400, height: 457_200 },
        textBody: { content },
      },
    },
  ],
});

const textBoxParagraph = (content: BlockContent[]): Paragraph => ({
  type: "paragraph",
  content: [textBoxRun(content)],
});

const observedPageBreakParagraph = (
  observation: SourceReadObservation,
  label: string,
): Paragraph => {
  const pageBreak: RunContent = observeTraversalRead({
    node: { type: "break", breakType: "page" },
    property: "breakType",
    label: `${label}:break`,
    observation,
  });
  const run = observeTraversalRead({
    node: { type: "run", content: [pageBreak] } satisfies Run,
    property: "content",
    label: `${label}:run`,
    observation,
    kind: "run",
  });
  return observeTraversalRead({
    node: { type: "paragraph", content: [run] } satisfies Paragraph,
    property: "content",
    label: `${label}:paragraph`,
    observation,
  });
};

const observedTextBoxParagraph = (
  content: BlockContent[],
  observation: SourceReadObservation,
  label: string,
): Paragraph => {
  const shapeContent: RunContent = observeTraversalRead({
    node: {
      type: "shape",
      shape: {
        type: "shape",
        shapeType: "rect",
        size: { width: 914_400, height: 457_200 },
        textBody: { content },
      },
    },
    property: "shape",
    label: `${label}:shape`,
    observation,
  });
  const run = observeTraversalRead({
    node: { type: "run", content: [shapeContent] } satisfies Run,
    property: "content",
    label: `${label}:run`,
    observation,
    kind: "run",
  });
  return observeTraversalRead({
    node: { type: "paragraph", content: [run] } satisfies Paragraph,
    property: "content",
    label: `${label}:paragraph`,
    observation,
  });
};

const observedTableWithCellContent = (
  content: BlockContent[],
  observation: SourceReadObservation,
  label: string,
): Table => {
  const cell = observeTraversalRead({
    node: { type: "tableCell", content } satisfies TableCell,
    property: "content",
    label: `${label}:cell`,
    observation,
  });
  const row = observeTraversalRead({
    node: { type: "tableRow", cells: [cell] } satisfies TableRow,
    property: "cells",
    label: `${label}:row`,
    observation,
  });
  return observeTraversalRead({
    node: { type: "table", rows: [row] } satisfies Table,
    property: "rows",
    label: `${label}:table`,
    observation,
  });
};

const nestedSource = (depth: number): NestedSource => {
  const ownerContents: BlockContent[][] = [];
  const observation: SourceReadObservation = {
    nodeReads: new Map(),
    runReads: new Map(),
  };
  let nested: BlockContent = observedPageBreakParagraph(observation, "root");
  let textBoxLayers = 0;

  for (let index = 0; index < depth; index += 1) {
    const content = [nested];
    ownerContents.push(content);
    if (index % 2 === 0) {
      nested = observedTableWithCellContent(content, observation, `layer-${String(index)}`);
      continue;
    }
    textBoxLayers += 1;
    nested = observedTextBoxParagraph(content, observation, `layer-${String(index)}`);
  }

  return {
    content: [nested],
    ownerContents,
    // The deepest break contributes a paragraph, run, and break node; every
    // table/text-box layer contributes three structurally observed nodes.
    expectedNodeReads: 3 + depth * 3,
    expectedRunReads: 1 + textBoxLayers,
    observation,
  };
};

const measureNestedSource = (depth: number) => {
  const source = nestedSource(depth);
  const index = buildPageBreakRunSourceDescendantIndex(source.content);

  expect(source.observation.nodeReads.size).toBe(source.expectedNodeReads);
  expect([...source.observation.nodeReads.values()].every((count) => count === 1)).toBe(true);
  expect(source.observation.runReads.size).toBe(source.expectedRunReads);
  expect([...source.observation.runReads.values()].every((count) => count === 1)).toBe(true);
  expect(source.ownerContents.every((content) => index.containsPageBreakRun(content))).toBe(true);
  return {
    nodeReads: source.observation.nodeReads.size,
    runReads: source.observation.runReads.size,
  };
};

describe("page-break source descendant index", () => {
  test("reads each source node and run traversal property once across doubled depths", () => {
    const depth = 128;
    const visitsAtDepth = measureNestedSource(depth);
    const visitsAtDoubleDepth = measureNestedSource(depth * 2);

    expect(visitsAtDoubleDepth.nodeReads - visitsAtDepth.nodeReads).toBe(depth * 3);
    expect(visitsAtDoubleDepth.runReads - visitsAtDepth.runReads).toBe(depth / 2);
  });

  test("indexes later and unrelated containers after finding an earlier break", () => {
    const firstContent: BlockContent[] = [pageBreakParagraph()];
    const unrelatedContent: BlockContent[] = [
      { type: "paragraph", content: [{ type: "run", content: [{ type: "text", text: "safe" }] }] },
    ];
    const laterContent: BlockContent[] = [pageBreakParagraph()];
    const content: BlockContent[] = [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: firstContent },
              { type: "tableCell", content: unrelatedContent },
            ],
          },
        ],
      },
      textBoxParagraph(laterContent),
    ];

    const index = buildPageBreakRunSourceDescendantIndex(content);

    expect(index.containsPageBreakRun(firstContent)).toBe(true);
    expect(index.containsPageBreakRun(unrelatedContent)).toBe(false);
    expect(index.containsPageBreakRun(laterContent)).toBe(true);
  });

  test("separates direct paragraph features from nested text-box content", () => {
    const separatedOwners: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "simpleField",
          instruction: "REF break",
          fieldType: "REF",
          content: [pageBreakRun()],
        },
        {
          type: "simpleField",
          instruction: "REF shape",
          fieldType: "REF",
          content: [textBoxRun([])],
        },
      ],
    };
    const sameRun: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "run",
          content: [{ type: "break", breakType: "page" }, ...textBoxRun([]).content],
        },
      ],
    };
    const nestedOnly = textBoxParagraph([pageBreakParagraph()]);
    const content = [separatedOwners, sameRun, nestedOnly];

    const index = buildPageBreakRunSourceDescendantIndex(content);

    expect(index.paragraphFeatures(separatedOwners)).toEqual({
      hasPageBreakRun: true,
      hasTextBoxShape: true,
      pageBreakSharesTextBoxShape: false,
    });
    expect(index.paragraphFeatures(sameRun)).toEqual({
      hasPageBreakRun: true,
      hasTextBoxShape: true,
      pageBreakSharesTextBoxShape: true,
    });
    expect(index.paragraphFeatures(nestedOnly)).toEqual({
      hasPageBreakRun: false,
      hasTextBoxShape: true,
      pageBreakSharesTextBoxShape: false,
    });
  });
});
