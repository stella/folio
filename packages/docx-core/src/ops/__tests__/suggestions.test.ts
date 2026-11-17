import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  TableCell,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  pendingSuggestionIdsIn,
  suggestionRevisionInfos,
  withoutPendingSuggestions,
} from "../suggestions";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const info = (id: number, suggestionId?: string) => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  ...(suggestionId === undefined ? {} : { suggestionId }),
});
const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000001", content }],
    },
  },
});
const contentOf = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

describe("a save without pending suggestions", () => {
  test("rejects every suggested revision and keeps ordinary tracked changes", () => {
    const tracked: ParagraphContent = {
      type: "insertion",
      info: info(3),
      content: [run(" kept")],
    };
    const saved = withoutPendingSuggestions(
      documentOf([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
        { type: "deletion", info: info(2, "s1"), content: [run(" world")] },
        tracked,
      ]),
    ).unwrap();
    expect(contentOf(saved)).toEqual([run("Hello world"), tracked]);
  });

  test("leaves a document without suggestions as it is", () => {
    const document = documentOf([
      run("Hello"),
      { type: "insertion", info: info(1), content: [run("!")] },
    ]);
    expect(withoutPendingSuggestions(document).unwrap()).toBe(document);
  });
});

describe("setting revision info", () => {
  const adopted = { id: 1, author: "Adopter", date: "2026-02-02T00:00:00Z" };
  const suggested = documentOf([
    run("Hello"),
    { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
    {
      type: "run",
      content: [{ type: "text", text: "!" }],
    },
    { type: "insertion", info: info(2, "s1"), content: [run("?")] },
  ]);
  const adopt = {
    type: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
    story: OP_STORIES.MAIN,
    infos: [adopted, { ...adopted, id: 2 }],
  } as const;

  test("writes each revision's info and undoes exactly", () => {
    const applied = applyDocumentOp(suggested, adopt).unwrap();
    expect(
      contentOf(applied.document).flatMap((item) =>
        item.type === "insertion" ? [item.info] : [],
      ),
    ).toEqual([adopted, { ...adopted, id: 2 }]);
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toEqual(suggested);
  });

  test("declares the paragraphs of the row, cell or table whose own revision it writes", () => {
    // A tracked row deletion as a package records it: the row's own
    // deletion, a deletion around the text it held, and property changes on
    // a cell and on the table.
    const paragraph = (
      paraId: string,
      content: ParagraphContent[],
    ): Paragraph => ({ type: "paragraph", paraId, content });
    const cell = (
      paraId: string,
      content: ParagraphContent[],
      extra: Partial<TableCell> = {},
    ): TableCell => ({
      ...extra,
      type: "tableCell",
      content: [paragraph(paraId, content)],
    });
    const tabled: Document = {
      package: {
        document: {
          content: [
            paragraph("00000001", [run("Before")]),
            {
              type: "table",
              propertyChanges: [
                {
                  type: "tablePropertyChange",
                  info: info(7),
                  previousFormatting: {},
                },
              ],
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    cell("00000002", [run("a")], {
                      propertyChanges: [
                        {
                          type: "tableCellPropertyChange",
                          info: info(8),
                          previousFormatting: {},
                        },
                      ],
                    }),
                    cell("00000003", [run("b")]),
                  ],
                },
                {
                  type: "tableRow",
                  structuralChange: { type: "tableRowDeletion", info: info(5) },
                  cells: [
                    cell("00000004", [
                      { type: "deletion", info: info(6), content: [run("c")] },
                    ]),
                    cell("00000005", [
                      { type: "deletion", info: info(9), content: [run("d")] },
                    ]),
                  ],
                },
              ],
            },
            paragraph("00000006", [run("After")]),
          ],
        },
      },
    };
    const declaredFor = (id: number) => {
      const applied = applyDocumentOp(tabled, {
        ...adopt,
        infos: [{ ...adopted, id }],
      }).unwrap();
      expect(
        applyDocumentOps(applied.document, applied.inverse).unwrap().document,
      ).toEqual(tabled);
      return applied.touched.modified.toSorted();
    };
    expect(declaredFor(5)).toEqual(["00000004", "00000005"]);
    expect(declaredFor(6)).toEqual(["00000004"]);
    expect(declaredFor(8)).toEqual(["00000002"]);
    expect(declaredFor(7)).toEqual([
      "00000002",
      "00000003",
      "00000004",
      "00000005",
    ]);
  });

  test("refuses an id the story does not hold", () => {
    expect(
      applyDocumentOp(suggested, {
        ...adopt,
        infos: [{ ...adopted, id: 9 }],
      }).isErr(),
    ).toBe(true);
  });
});

describe("final section suggestion boundaries", () => {
  const sectionChange = (suggestionId?: string) => {
    const source = normalizeForOps(
      documentOf([
        run("Standing"),
        {
          type: "insertion",
          info: info(1, suggestionId),
          content: [run(" added")],
        },
      ]),
    );
    return applyDocumentOp(source, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      sectionIndex: 0,
      patch: {
        pageWidth: 14400,
        propertyChanges: [
          {
            type: "sectionPropertyChange",
            info: info(100, suggestionId),
            previousProperties: { pageWidth: 12240 },
          },
        ],
      },
    }).unwrap().document;
  };

  test("finds both source revisions of a proposal made by section operations", () => {
    const document = sectionChange("proposal");
    expect([...pendingSuggestionIdsIn(document)]).toEqual(["proposal"]);
    expect(
      suggestionRevisionInfos(document, OP_STORIES.MAIN, "proposal"),
    ).toEqual([info(1, "proposal"), info(100, "proposal")]);
    expect(suggestionRevisionInfos(document, OP_STORIES.MAIN, "other")).toEqual(
      [],
    );
  });

  test.each(Object.values(REVISION_DECISIONS))(
    "%s refuses a final section revision before resolving any body revision",
    (decision) => {
      const document = sectionChange("proposal");
      const before = structuredClone(document);
      for (const revisionIds of [[100], [1, 100], [100, 1]]) {
        const result = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds,
          decision,
        });
        expect(result.isErr()).toBe(true);
        if (result.isOk()) panic("Expected section review refusal");
        expect(result.error.reason).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        );
        expect(document).toEqual(before);
      }
    },
  );

  test("a save refuses instead of leaving a pending section proposal in the file", () => {
    const document = sectionChange("proposal");
    const before = structuredClone(document);
    const result = withoutPendingSuggestions(document);
    expect(result.isErr()).toBe(true);
    if (result.isOk()) panic("Expected section review refusal");
    expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    expect(document).toEqual(before);
  });

  test.each(["source-only", "with-derived-view"])(
    "adoption refuses atomically with %s",
    (view) => {
      const document = sectionChange("proposal");
      if (view === "with-derived-view") {
        const body = document.package.document;
        if (body.finalSectionProperties === undefined)
          panic("The producer writes section properties");
        body.sections = [
          { properties: body.finalSectionProperties, content: body.content },
        ];
      }
      const before = structuredClone(document);
      const result = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
        story: OP_STORIES.MAIN,
        infos: [info(1), info(100)],
      });
      expect(result.isErr()).toBe(true);
      if (result.isOk()) panic("Expected section review refusal");
      expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
      expect(document).toEqual(before);
    },
  );

  test("ordinary section changes remain saveable and do not block body-only review", () => {
    const document = sectionChange();
    expect([...pendingSuggestionIdsIn(document)]).toEqual([]);
    expect(withoutPendingSuggestions(document).unwrap()).toBe(document);
    const resolved = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [1],
      decision: REVISION_DECISIONS.ACCEPT,
    }).unwrap();
    expect(contentOf(resolved.document)).toEqual([run("Standing added")]);
    expect(resolved.document.package.document.finalSectionProperties).toEqual(
      document.package.document.finalSectionProperties,
    );
    expect(
      applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
    ).toEqual(document);
  });
});
