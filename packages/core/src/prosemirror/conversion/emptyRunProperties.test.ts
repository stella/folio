import { expect, test } from "bun:test";

import type { Document, Run } from "../../types/document";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const roundTrip = (runs: Run[]): Run[] => {
  const source: Document = {
    package: { document: { content: [{ type: "paragraph", content: runs }] } },
  };
  const saved = fromProseDoc(toProseDoc(source), source);
  const paragraph = saved.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected one paragraph");
  }
  return paragraph.content.filter((item): item is Run => item.type === "run");
};

test("an authored empty run-property element survives the editor round trip", () => {
  const runs: Run[] = [
    { type: "run", content: [{ type: "text", text: "before" }], formatting: {} },
    { type: "run", content: [{ type: "text", text: "after" }] },
  ];
  expect(roundTrip(runs)).toEqual(runs);
  expect(roundTrip(roundTrip(runs))).toEqual(runs);
});

test("empty run properties survive when a run also contains an inline atom", () => {
  const runs: Run[] = [
    {
      type: "run",
      content: [{ type: "text", text: "before" }, { type: "tab" }],
      formatting: {},
    },
  ];
  expect(roundTrip(runs)).toEqual(runs);
});
