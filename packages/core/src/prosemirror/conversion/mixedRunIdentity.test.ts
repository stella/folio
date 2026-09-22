/** One authored run containing an inline atom stays one run after an editor save. */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent, Run } from "../../types/document";

import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const CONTENTS = [
  [{ type: "tab" }, { type: "text", text: "after" }],
  [{ type: "text", text: "before" }, { type: "tab" }],
  [{ type: "text", text: "before" }, { type: "tab" }, { type: "text", text: "after" }],
  [
    { type: "text", text: "before" },
    { type: "break", breakType: "textWrapping" },
  ],
  [
    { type: "break", breakType: "textWrapping" },
    { type: "text", text: "after" },
  ],
  [{ type: "tab" }, { type: "tab" }],
] as const;

const HOSTS = ["paragraph", "hyperlink", "insertion"] as const;

const wrap = (host: (typeof HOSTS)[number], run: Run): ParagraphContent => {
  switch (host) {
    case "paragraph":
      return run;
    case "hyperlink":
      return { type: "hyperlink", href: "https://example.test", children: [run] };
    case "insertion":
      return { type: "insertion", info: { id: 7, author: "Reviewer" }, content: [run] };
  }
};

const innerRuns = (item: ParagraphContent): Run[] => {
  switch (item.type) {
    case "run":
      return [item];
    case "hyperlink":
      return item.children.filter((child): child is Run => child.type === "run");
    case "insertion":
      return item.content.filter((child): child is Run => child.type === "run");
    default:
      throw new Error(`Unexpected paragraph content: ${item.type}`);
  }
};

for (const host of HOSTS) {
  for (const [index, contents] of CONTENTS.entries()) {
    test(`${host} keeps one mixed run in position ${index}`, () => {
      const run: Run = { type: "run", content: [...contents] };
      const document: Document = {
        package: { document: { content: [{ type: "paragraph", content: [wrap(host, run)] }] } },
      };
      const saved = fromProseDoc(toProseDoc(document), document);
      const paragraph = saved.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected one paragraph");
      }
      expect(paragraph.content).toHaveLength(1);
      const item = paragraph.content.at(0);
      if (!item) {
        throw new Error("Expected paragraph content");
      }
      expect(item.type).toBe(host === "paragraph" ? "run" : host);
      expect(innerRuns(item).map((part) => part.content)).toEqual([run.content]);

      const savedAgain = fromProseDoc(toProseDoc(saved), saved);
      const again = savedAgain.package.document.content.at(0);
      expect(again?.type === "paragraph" ? again.content : []).toEqual(paragraph.content);
    });
  }
}

test("adjacent mixed runs retain distinct authored boundaries", () => {
  const first: Run = {
    type: "run",
    content: [{ type: "text", text: "first" }, { type: "tab" }],
  };
  const second: Run = {
    type: "run",
    content: [{ type: "tab" }, { type: "text", text: "second" }],
  };
  const document: Document = {
    package: { document: { content: [{ type: "paragraph", content: [first, second] }] } },
  };
  const saved = fromProseDoc(toProseDoc(document), document);
  const paragraph = saved.package.document.content.at(0);
  expect(paragraph?.type === "paragraph" ? paragraph.content : []).toEqual([first, second]);
});

test("an inline atom and text retain their shared direct formatting", () => {
  const run: Run = {
    type: "run",
    content: [{ type: "tab" }, { type: "text", text: "label" }],
    formatting: { bold: true },
  };
  const document: Document = {
    package: { document: { content: [{ type: "paragraph", content: [run] }] } },
  };
  const saved = fromProseDoc(toProseDoc(document), document);
  const paragraph = saved.package.document.content.at(0);
  expect(paragraph?.type === "paragraph" ? paragraph.content : []).toEqual([run]);
});
