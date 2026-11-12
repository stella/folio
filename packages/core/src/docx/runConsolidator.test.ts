import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import type { Hyperlink, Run } from "../types/document";
import { parseDocumentBody } from "./documentParser";
import {
  consolidateParagraphContent,
  consolidateRuns,
} from "./runConsolidator";
import { serializeParagraph } from "./serializer/paragraphSerializer";

const plainRun = (text: string): Run => ({
  type: "run",
  formatting: { fontSizeCs: 22 },
  content: [{ type: "text", text }],
});

const changedRun = (text: string, id = 1): Run => ({
  type: "run",
  formatting: { fontSizeCs: 22 },
  propertyChanges: [
    {
      type: "runPropertyChange",
      info: {
        id,
        author: "Reviewer",
        date: "2026-09-08T08:00:00Z",
      },
      previousFormatting: { bold: true, fontSizeCs: 22 },
      currentFormatting: { fontSizeCs: 22 },
    },
  ],
  content: [{ type: "text", text }],
});

const runText = (run: Run): string =>
  run.content
    .map((content) => (content.type === "text" ? content.text : ""))
    .join("");

describe("run consolidation boundaries", () => {
  test.each([
    {
      name: "at the start",
      runs: [changedRun("changed"), plainRun("after"), plainRun("end")],
      expectedText: ["changed", "afterend"],
      expectedChangedIndex: 0,
    },
    {
      name: "in the middle",
      runs: [plainRun("before"), changedRun("changed"), plainRun("after")],
      expectedText: ["before", "changed", "after"],
      expectedChangedIndex: 1,
    },
    {
      name: "at the end",
      runs: [plainRun("before"), plainRun("middle"), changedRun("changed")],
      expectedText: ["beforemiddle", "changed"],
      expectedChangedIndex: 1,
    },
  ])(
    "keeps a run-property revision range $name",
    ({ runs, expectedText, expectedChangedIndex }) => {
      const result = consolidateRuns(runs);

      expect(result.map(runText)).toEqual(expectedText);
      expect(result.at(expectedChangedIndex)).toEqual(changedRun("changed"));
      expect(
        result.filter((run) => run.propertyChanges?.length).map(runText),
      ).toEqual(["changed"]);
    },
  );

  test("keeps adjacent run-property revisions separate", () => {
    const result = consolidateRuns([
      plainRun("before"),
      changedRun("first", 1),
      changedRun("second", 2),
      plainRun("after"),
    ]);

    expect(result.map(runText)).toEqual(["before", "first", "second", "after"]);
    expect(result.map((run) => run.propertyChanges?.at(0)?.info.id)).toEqual([
      undefined,
      1,
      2,
      undefined,
    ]);
  });

  test("keeps a run-property revision range inside a hyperlink", () => {
    const hyperlink: Hyperlink = {
      type: "hyperlink",
      href: "https://example.test",
      children: [plainRun("before"), changedRun("changed"), plainRun("after")],
    };

    const result = consolidateParagraphContent([hyperlink]);

    expect(result).toHaveLength(1);
    expect(result.at(0)).toEqual({
      ...hyperlink,
      children: [plainRun("before"), changedRun("changed"), plainRun("after")],
    });
  });

  test("merges runs when their property-change collection is empty", () => {
    const result = consolidateRuns([
      { ...plainRun("before"), propertyChanges: [] },
      { ...plainRun("after"), propertyChanges: [] },
    ]);

    expect(result).toHaveLength(1);
    expect(result.map(runText)).toEqual(["beforeafter"]);
  });

  test("keeps the changed text in its own serialized run after parsing", () => {
    const body = parseDocumentBody(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r><w:rPr><w:szCs w:val="22"/></w:rPr><w:t>before</w:t></w:r>
            <w:r><w:rPr><w:szCs w:val="22"/><w:rPrChange w:id="1" w:author="Reviewer" w:date="2026-09-08T08:00:00Z"><w:rPr><w:b/><w:szCs w:val="22"/></w:rPr></w:rPrChange></w:rPr><w:t>changed</w:t></w:r>
            <w:r><w:rPr><w:szCs w:val="22"/></w:rPr><w:t>after</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const paragraph = body.content.at(0);

    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      panic("Expected parsed paragraph");
    }

    expect(
      paragraph.content
        .filter((content) => content.type === "run")
        .map(runText),
    ).toEqual(["before", "changed", "after"]);

    const xml = serializeParagraph(paragraph);
    expect(xml).toMatch(
      /<w:r>[^]*?<w:t>before<\/w:t><\/w:r><w:r><w:rPr>[^]*?<w:rPrChange\b[^]*?<\/w:rPrChange><\/w:rPr><w:t>changed<\/w:t><\/w:r><w:r>[^]*?<w:t>after<\/w:t><\/w:r>/u,
    );
  });
});

describe("tabs and line breaks flow with text", () => {
  const runOf = (content: Run["content"]): Run => ({
    type: "run",
    formatting: { fontSizeCs: 22 },
    content,
  });

  test("alike runs split beside a tab or a line break read as one run", () => {
    // A save writes the pieces an edit cut from one run as runs of their own;
    // read back, they are the run the editor held.
    const pieces = [
      runOf([{ type: "text", text: "Delivery" }]),
      runOf([{ type: "break" }]),
      runOf([{ type: "text", text: "Payment" }, { type: "tab" }]),
      runOf([
        { type: "break", sourceElement: "cr" },
        { type: "text", text: "Notice" },
      ]),
      runOf([{ type: "break", breakType: "textWrapping" }]),
    ];
    expect(consolidateRuns(pieces)).toEqual([
      runOf([
        { type: "text", text: "Delivery" },
        { type: "break" },
        { type: "text", text: "Payment" },
        { type: "tab" },
        { type: "break", sourceElement: "cr" },
        { type: "text", text: "Notice" },
        { type: "break", breakType: "textWrapping" },
      ]),
    ]);
  });

  test("a page or column break and a positional tab stay merge boundaries", () => {
    const boundaries: Run["content"] = [
      { type: "break", breakType: "page" },
      { type: "break", breakType: "column" },
      {
        type: "tab",
        positional: {
          alignment: "right",
          relativeTo: "margin",
          leader: "none",
        },
      },
    ];
    for (const boundary of boundaries) {
      const runs = [
        runOf([{ type: "text", text: "a" }]),
        runOf([boundary]),
        runOf([{ type: "text", text: "b" }]),
      ];
      expect(consolidateRuns(runs)).toEqual(runs);
    }
  });
});
