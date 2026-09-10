import { describe, expect, spyOn, test } from "bun:test";
import { Node as PMNode } from "prosemirror-model";

import { schema } from "../prosemirror/schema";
import { buildPageBreakRunDescendantIndex } from "./pageBreakRunDescendantIndex";

type NestedPageBreakDocument = {
  doc: PMNode;
  owners: readonly PMNode[];
  pageBreak: PMNode;
  unrelatedParagraph: PMNode;
};

const nestedPageBreakDocument = (depth: number): NestedPageBreakDocument => {
  const pageBreak = schema.node("pageBreakRun");
  const deepestParagraph = schema.node("paragraph", null, [pageBreak]);
  const owners: PMNode[] = [deepestParagraph];
  let nestedBlock = deepestParagraph;

  for (let index = 0; index < depth; index += 1) {
    const cell = schema.node(index % 2 === 0 ? "tableCell" : "tableHeader", null, [nestedBlock]);
    owners.push(cell);
    nestedBlock = schema.node("table", null, [schema.node("tableRow", null, [cell])]);
  }

  const textBox = schema.node("textBox", { width: 100 }, [nestedBlock]);
  owners.push(textBox);
  const unrelatedParagraph = schema.node("paragraph", null, [schema.text("sibling")]);
  return {
    doc: schema.node("doc", null, [textBox, unrelatedParagraph]),
    owners,
    pageBreak,
    unrelatedParagraph,
  };
};

const absolutePositionOf = (doc: PMNode, target: PMNode): number => {
  let position: number | undefined;
  doc.descendants((node, nodePosition) => {
    if (node === target) {
      position = nodePosition;
    }
  });
  if (position === undefined) {
    throw new Error("Expected target node in test document");
  }
  return position;
};

const measureIndexBuild = (depth: number): number => {
  const { doc, owners, pageBreak, unrelatedParagraph } = nestedPageBreakDocument(depth);
  const { index, visitCount } = (() => {
    const forEach = spyOn(PMNode.prototype, "forEach");
    try {
      return {
        index: buildPageBreakRunDescendantIndex(doc),
        visitCount: forEach.mock.calls.length,
      };
    } finally {
      forEach.mockRestore();
    }
  })();

  let nodeCount = 1;
  doc.descendants(() => {
    nodeCount += 1;
  });
  expect(visitCount).toBe(nodeCount);

  const pageBreakPosition = absolutePositionOf(doc, pageBreak);
  expect(
    owners.every((owner) => index.firstPageBreakRunPosition(owner) === pageBreakPosition),
  ).toBe(true);
  expect(index.firstPageBreakRunPosition(unrelatedParagraph)).toBeUndefined();
  return visitCount;
};

describe("page-break descendant index", () => {
  test("visits every node exactly once across adversarial nested containers", () => {
    const depth = 128;
    const visitsAtDepth = measureIndexBuild(depth);
    const visitsAtDoubleDepth = measureIndexBuild(depth * 2);

    expect(visitsAtDoubleDepth - visitsAtDepth).toBe(depth * 3);
  });

  test("indexes later container children after an earlier break is found", () => {
    const firstBreak = schema.node("pageBreakRun");
    const firstParagraph = schema.node("paragraph", null, [firstBreak]);
    const laterBreak = schema.node("pageBreakRun");
    const laterParagraph = schema.node("paragraph", null, [laterBreak]);
    const laterCell = schema.node("tableCell", null, [laterParagraph]);
    const laterTable = schema.node("table", null, [schema.node("tableRow", null, [laterCell])]);
    const textBox = schema.node("textBox", { width: 100 }, [firstParagraph, laterTable]);
    const doc = schema.node("doc", null, [textBox]);

    const index = buildPageBreakRunDescendantIndex(doc);

    expect(index.firstPageBreakRunPosition(textBox)).toBe(absolutePositionOf(doc, firstBreak));
    expect(index.firstPageBreakRunPosition(laterCell)).toBe(absolutePositionOf(doc, laterBreak));
    expect(index.firstPageBreakRunPosition(laterParagraph)).toBe(
      absolutePositionOf(doc, laterBreak),
    );
  });
});
