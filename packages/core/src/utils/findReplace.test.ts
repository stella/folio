import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import type { Document } from "../types/document";
import type { HeaderFooter, Paragraph, Run } from "../types/content";
import { FindReplaceManager } from "../managers/FindReplaceManager";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { findInProseMirrorDocument } from "../prosemirror/findReplaceSelection";
import { createDefaultFindOptions, findAllMatches, findInDocument } from "./findReplace";

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Outside stock" }] }],
        },
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [
                    {
                      type: "paragraph",
                      content: [{ type: "run", content: [{ type: "text", text: "Inside stock" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
};

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const allStoriesDocument = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("body hit"),
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    size: { width: 914400, height: 914400 },
                    textBody: { content: [paragraph("box hit")] },
                  },
                },
              ],
            },
          ],
        },
        paragraph("last body hit"),
      ],
      sections: [
        {
          properties: {
            headerReferences: [
              { type: "even", rId: "he" },
              { type: "first", rId: "hf" },
              { type: "default", rId: "hd" },
            ],
            footerReferences: [
              { type: "even", rId: "fe" },
              { type: "first", rId: "ff" },
              { type: "default", rId: "fd" },
            ],
          },
          content: [],
        },
      ],
    },
    headers: new Map(
      ["he", "hf", "hd"].map((rId) => [
        rId,
        {
          type: "header",
          hdrFtrType: "default",
          content: [paragraph(`${rId} hit`)],
        } as const satisfies HeaderFooter,
      ]),
    ),
    footers: new Map(
      ["fe", "ff", "fd"].map((rId) => [
        rId,
        {
          type: "footer",
          hdrFtrType: "default",
          content: [paragraph(`${rId} hit`)],
        } as const satisfies HeaderFooter,
      ]),
    ),
    footnotes: [
      { type: "footnote", id: -1, noteType: "separator", content: [paragraph("separator hit")] },
      { type: "footnote", id: 1, content: [paragraph("footnote hit")] },
    ],
    endnotes: [{ type: "endnote", id: 2, content: [paragraph("endnote hit")] }],
  },
});

describe("document-wide story search", () => {
  test("covers each story and variant once, with stable navigation order and shape addresses", () => {
    const doc = allStoriesDocument();
    const before = structuredClone(doc);
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    expect(matches).toHaveLength(11);
    expect(matches.map(({ story, textBoxIndex }) => ({ story, textBoxIndex }))).toEqual([
      { story: "main", textBoxIndex: undefined },
      { story: "main", textBoxIndex: 0 },
      { story: "main", textBoxIndex: undefined },
      ...["hd", "hf", "he"].map((rId) => ({
        story: { kind: "header", rId },
        textBoxIndex: undefined,
      })),
      ...["fd", "ff", "fe"].map((rId) => ({
        story: { kind: "footer", rId },
        textBoxIndex: undefined,
      })),
      { story: { kind: "footnote", id: 1 }, textBoxIndex: undefined },
      { story: { kind: "endnote", id: 2 }, textBoxIndex: undefined },
    ]);
    expect(matches.at(1)).toMatchObject({
      paragraphIndex: 0,
      anchorParagraphIndex: 1,
      textBoxIndex: 0,
    });
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    for (const match of matches.slice(1)) expect(manager.navigate("next")?.match).toEqual(match);
    expect(manager.navigate("next")?.match).toEqual(matches.at(0));
    expect(manager.navigate("previous")?.match).toEqual(matches.at(-1));
    expect(doc).toEqual(before);
  });

  test.each(["all-markup", "simple-markup", "no-markup", "original"] as const)(
    "%s respects revision visibility without shifting authored offsets",
    (displayMode) => {
      const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });
      const doc: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  run("before "),
                  {
                    type: "deletion",
                    info: { id: 1, author: "Author" },
                    content: [run("deleted hit ")],
                  },
                  {
                    type: "insertion",
                    info: { id: 2, author: "Author" },
                    content: [run("inserted hit ")],
                  },
                  run("after hit"),
                ],
              },
            ],
          },
        },
      };
      const options = { ...createDefaultFindOptions(), displayMode };
      const matches = findInDocument(doc, "hit", options);
      expect(matches).toHaveLength(displayMode === "all-markup" ? 3 : 2);
      expect(matches.at(-1)?.startOffset).toBe(38);
      expect(findInDocument(doc, "deleted", options)).toHaveLength(
        displayMode === "original" || displayMode === "all-markup" ? 1 : 0,
      );
      expect(findInDocument(doc, "inserted", options)).toHaveLength(
        displayMode === "original" ? 0 : 1,
      );
    },
  );

  test("secondary-story replacement cannot target a coincident body paragraph", () => {
    const doc = allStoriesDocument();
    const matches = findInDocument(doc, "hit", createDefaultFindOptions());
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    for (const [index, match] of matches.entries()) {
      if (match.story === "main" && match.textBoxIndex === undefined) continue;
      expect(manager.goTo(index)?.match).toEqual(match);
      expect(manager.replaceCurrent(doc, "changed")).toBeNull();
    }
    expect(manager.replaceAll(doc, matches, "changed")).toBeNull();
    expect(findInDocument(doc, "hit", createDefaultFindOptions())).toHaveLength(11);
  });

  test("only rendered variants count when the view supplies story eligibility", () => {
    const matches = findInDocument(allStoriesDocument(), "hit", createDefaultFindOptions(), {
      renderedStories: [
        "main",
        { kind: "header", rId: "hd" },
        { kind: "header", rId: "hd" },
        { kind: "footer", rId: "fd" },
      ],
    });
    expect(matches).toHaveLength(7);
    expect(matches.map(({ story }) => story)).toEqual([
      "main",
      "main",
      "main",
      { kind: "header", rId: "hd" },
      { kind: "footer", rId: "fd" },
      { kind: "footnote", id: 1 },
      { kind: "endnote", id: 2 },
    ]);
  });

  test("every story keeps case and whole-word behavior", () => {
    const doc = allStoriesDocument();
    expect(findInDocument(doc, "HIT", { matchCase: false, matchWholeWord: true })).toHaveLength(11);
    expect(findInDocument(doc, "HIT", { matchCase: true, matchWholeWord: true })).toHaveLength(0);
    expect(findInDocument(doc, "hi", { matchCase: false, matchWholeWord: true })).toHaveLength(0);
    expect(findInDocument(doc, "hi", { matchCase: false, matchWholeWord: false })).toHaveLength(11);
  });

  test("hidden shapes retain authored numbering for later visible shape addresses", () => {
    const doc = allStoriesDocument();
    const shapeAnchor = doc.package.document.content.at(1);
    if (shapeAnchor?.type !== "paragraph") panic("Fixture shape anchor missing");
    const shapeRun = shapeAnchor.content.at(0);
    if (shapeRun?.type !== "run") panic("Fixture shape run missing");
    doc.package.document.content = [
      {
        type: "paragraph",
        content: [
          { type: "deletion", info: { id: 1, author: "Author" }, content: [shapeRun] },
          shapeRun,
        ],
      },
    ];
    const matches = findInDocument(doc, "box", {
      ...createDefaultFindOptions(),
      displayMode: "no-markup",
    });
    expect(matches).toMatchObject([
      { story: "main", textBoxIndex: 1, anchorParagraphIndex: 0, paragraphIndex: 0 },
    ]);
  });

  test("comment text is opt-in, has its own address, and cannot replace body text", () => {
    const doc = allStoriesDocument();
    doc.package.document.comments = [
      { id: 7, author: "Author", content: [paragraph("comment hit")] },
    ];
    expect(findInDocument(doc, "comment", createDefaultFindOptions())).toHaveLength(0);
    const matches = findInDocument(doc, "comment", {
      ...createDefaultFindOptions(),
      includeComments: true,
    });
    expect(matches).toMatchObject([
      { story: "main", commentId: 7, paragraphIndex: 0, startOffset: 0 },
    ]);
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    expect(manager.replaceCurrent(doc, "changed")).toBeNull();
    expect(manager.replaceAll(doc, matches, "changed")).toBeNull();
  });

  test.each(["all-markup", "simple-markup", "no-markup", "original"] as const)(
    "%s respects table row and cell visibility and keeps authored paragraph indices",
    (displayMode) => {
      const doc: Document = {
        package: {
          document: {
            content: [
              {
                type: "table",
                rows: [
                  {
                    type: "tableRow",
                    structuralChange: {
                      type: "tableRowDeletion",
                      info: { id: 1, author: "Author" },
                    },
                    cells: [{ type: "tableCell", content: [paragraph("deleted row hit")] }],
                  },
                  {
                    type: "tableRow",
                    cells: [
                      {
                        type: "tableCell",
                        structuralChange: {
                          type: "tableCellInsertion",
                          info: { id: 2, author: "Author" },
                        },
                        content: [paragraph("inserted cell hit")],
                      },
                      { type: "tableCell", content: [paragraph("visible hit")] },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const matches = findInDocument(doc, "hit", { ...createDefaultFindOptions(), displayMode });
      const expected = {
        "all-markup": [0, 1, 2],
        original: [0, 2],
        "no-markup": [1, 2],
        "simple-markup": [1, 2],
      };
      expect(matches.map(({ paragraphIndex }) => paragraphIndex)).toEqual(expected[displayMode]);
    },
  );
});

describe("regex matching", () => {
  test("honors regex mode and groups whole-word alternation", () => {
    expect(
      findAllMatches("hit 42 hit 99", "hit \\d+", {
        matchCase: true,
        matchWholeWord: false,
        useRegex: true,
      }),
    ).toEqual([
      { start: 0, end: 6 },
      { start: 7, end: 13 },
    ]);
    expect(
      findAllMatches("foox xbar foo bar", "foo|bar", {
        matchCase: false,
        matchWholeWord: true,
        useRegex: true,
      }),
    ).toEqual([
      { start: 10, end: 13 },
      { start: 14, end: 17 },
    ]);
    expect(
      findAllMatches("hit", "[", { matchCase: false, matchWholeWord: false, useRegex: true }),
    ).toEqual([]);
  });

  test("skips empty regex hits and advances over complete Unicode code points", () => {
    const options = { matchCase: false, matchWholeWord: false, useRegex: true };
    expect(findAllMatches("📝 hit 📝 hit", "(?=📝)|hit", options)).toEqual([
      { start: 3, end: 6 },
      { start: 10, end: 13 },
    ]);
    for (const pattern of ["^", "$", "a*", "(?=.)", "(?:)"]) {
      expect(findAllMatches("📝 xyz", pattern, options)).toEqual([]);
    }
  });
});

describe("findInDocument", () => {
  test("finds matches in body and table paragraphs with stable paragraph offsets", () => {
    const matches = findInDocument(document, "stock", createDefaultFindOptions());

    expect(
      matches.map(({ paragraphIndex, startOffset, text }) => ({
        paragraphIndex,
        startOffset,
        text,
      })),
    ).toEqual([
      { paragraphIndex: 0, startOffset: 8, text: "stock" },
      { paragraphIndex: 1, startOffset: 7, text: "stock" },
    ]);
  });

  test("honors case and whole-word options", () => {
    const caseDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Stock stock stockholder" }],
                },
              ],
            },
          ],
        },
      },
    };

    expect(
      findInDocument(caseDocument, "Stock", { matchCase: true, matchWholeWord: true }),
    ).toHaveLength(1);
    expect(
      findInDocument(caseDocument, "stock", { matchCase: false, matchWholeWord: true }),
    ).toHaveLength(2);
  });

  test("symbol run children contribute visible characters and preserve later offsets", () => {
    const doc: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text: "A" },
                    { type: "symbol", char: "00A7", font: "Arial" },
                    { type: "text", text: "B" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };
    expect(findInDocument(doc, "A§B", createDefaultFindOptions())).toMatchObject([
      { startOffset: 0, endOffset: 3 },
    ]);
    expect(findInDocument(doc, "B", createDefaultFindOptions())).toMatchObject([
      { startOffset: 2, endOffset: 3 },
    ]);
    const pmDoc = toProseDoc(doc);
    for (const query of ["A§B", "§", "B"]) {
      const modelMatches = findInDocument(doc, query, createDefaultFindOptions());
      const pmMatches = findInProseMirrorDocument(pmDoc, query, createDefaultFindOptions());
      expect(
        modelMatches.map(({ paragraphIndex, startOffset, endOffset, text }) => ({
          paragraphIndex,
          startOffset,
          endOffset,
          text,
        })),
      ).toEqual(
        pmMatches.map(({ paragraphIndex, startOffset, endOffset, text }) => ({
          paragraphIndex,
          startOffset,
          endOffset,
          text,
        })),
      );
    }
  });

  test("does not match text across an explicit page-break run", () => {
    const pageBreakDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text: "left" },
                    { type: "break", breakType: "page" },
                    { type: "text", text: "right" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    expect(findInDocument(pageBreakDocument, "left", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "right", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "leftright", createDefaultFindOptions())).toEqual([]);
  });

  test.each([undefined, "textWrapping"] as const)(
    "counts a %s line-break type in search offsets",
    (breakType) => {
      const lineBreakDocument: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "A" },
                      {
                        type: "break",
                        ...(breakType !== undefined ? { breakType } : {}),
                      },
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };

      expect(findInDocument(lineBreakDocument, "A\nB", createDefaultFindOptions())).toHaveLength(1);
      expect(
        findInDocument(lineBreakDocument, "B", createDefaultFindOptions()).at(0),
      ).toMatchObject({ startOffset: 2 });
    },
  );
});
