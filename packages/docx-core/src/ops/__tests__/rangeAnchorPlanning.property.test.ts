import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent, paragraphVisibleText } from "../editorIntent";
import { isRangeBoundary } from "../offsets";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";
import { RANGE_ANCHOR_FIXTURE_FACTORIES } from "../../../typecheck/range-anchor-fixtures.typecheck";

setDefaultTimeout(propertyTestTimeout(60000));

const run = (text: string) => ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
const PAIRS = Object.values(RANGE_ANCHOR_FIXTURE_FACTORIES).map((factory) => factory());
type RangePair = (typeof PAIRS)[number];

const paragraphOf = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Planned replacement lost its paragraph.");
  return paragraph;
};
const documentFor = (paragraph: Paragraph, pair: RangePair) =>
  ({
    package: {
      document: {
        content: [paragraph],
        ...(pair[0].type === "commentRangeStart"
          ? {
              comments: [
                {
                  id: pair[0].id,
                  author: "Reviewer",
                  content: [{ type: "paragraph", paraId: "22345678", content: [run("Comment")] }],
                },
              ],
            }
          : {}),
      },
    },
  }) satisfies Document;
const markersOf = (document: Document) => paragraphOf(document).content.filter(isRangeBoundary);

type TextOptions = { text: string; ownership: "existing" | "own" | "other"; id: number };
const sourceText = ({ text, ownership, id }: TextOptions): ParagraphContent =>
  ownership === "existing"
    ? run(text)
    : {
        type: "insertion",
        info: { id, author: ownership === "own" ? "Reviewer" : "Other" },
        content: [run(text)],
      };

test("replacement plans preserve range metadata across grapheme endpoints and ownership", () => {
  const check = ({
    pair,
    ownership,
    placement,
    base,
    suffix,
    replacement,
  }: {
    pair: RangePair;
    ownership: TextOptions["ownership"];
    placement: "internal" | "around";
    base: string;
    suffix: string;
    replacement: string;
  }) => {
    const grapheme = base + suffix;
    const selected =
      placement === "internal"
        ? [
            sourceText({ text: base, ownership, id: 21 }),
            pair[0],
            sourceText({ text: suffix, ownership, id: 22 }),
            pair[1],
          ]
        : [pair[0], sourceText({ text: grapheme, ownership, id: 21 }), pair[1]];
    const paragraph = {
      type: "paragraph",
      paraId: "12345678",
      content: [run("L"), ...selected, run("R")],
    } satisfies Paragraph;
    const document = documentFor(paragraph, pair);
    const from = {
      story: OP_STORIES.MAIN,
      blockId: paragraph.paraId,
      offset: 1,
      zeroWidthBefore: 0,
    } as const;
    const to = {
      story: OP_STORIES.MAIN,
      blockId: paragraph.paraId,
      offset: 1 + grapheme.length,
      zeroWidthBefore: placement === "around" ? 1 : 0,
    } as const;
    const newIds = { revision: Array.from({ length: 32 }, (_, index) => 1001 + index) };
    for (const mode of [
      { type: "editing", newIds },
      { type: "suggesting", revision: { id: 1000, author: "Reviewer" }, newIds },
    ] as const) {
      const compiled = compileEditorIntent(document, {
        intent: { type: "replaceText", from, to, text: replacement },
        mode,
      }).unwrap();
      const applied = applyDocumentOps(document, compiled.ops).unwrap();
      expect(paragraphVisibleText(paragraphOf(applied.document))).toBe(`L${replacement}R`);
      expect(markersOf(applied.document)).toEqual([...pair]);
      expect(applyDocumentOps(applied.document, applied.inverse).unwrap().document).toStrictEqual(
        document,
      );
      for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
        const resolved =
          applied.revisions.length === 0
            ? applied.document
            : applyDocumentOps(applied.document, [
                {
                  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                  story: OP_STORIES.MAIN,
                  revisionIds: applied.revisions,
                  decision,
                },
              ]).unwrap().document;
        const restoredText =
          mode.type === "suggesting" &&
          ownership !== "own" &&
          decision === REVISION_DECISIONS.REJECT;
        let expectedText = `L${replacement}R`;
        if (decision === REVISION_DECISIONS.REJECT && mode.type === "suggesting")
          expectedText = restoredText ? `L${grapheme}R` : "LR";
        expect(paragraphVisibleText(paragraphOf(resolved))).toBe(expectedText);
        expect(markersOf(resolved)).toEqual([...pair]);
      }
    }
    if (placement === "around" && replacement === "") {
      const raw = applyDocumentOps(document, [
        { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from, to },
      ]).unwrap();
      expect(markersOf(raw.document)).toEqual([]);
      expect(paragraphVisibleText(paragraphOf(raw.document))).toBe("LR");
      expect(applyDocumentOps(raw.document, raw.inverse).unwrap().document).toStrictEqual(document);
    }
  };
  for (const pair of PAIRS)
    for (const ownership of ["existing", "own", "other"] as const)
      for (const placement of ["internal", "around"] as const)
        for (const replacement of ["", "x", "契"])
          check({ pair, ownership, placement, base: "e", suffix: "\u0301", replacement });
  assertProperty(
    fc.property(
      fc.record({
        pair: fc.constantFrom(...PAIRS),
        ownership: fc.constantFrom("existing", "own", "other"),
        placement: fc.constantFrom("internal", "around"),
        base: fc.constantFrom("e", "a", "n"),
        suffix: fc.constantFrom("\u0301", "\u0308", "\u0303"),
        replacement: fc.constantFrom("", "x", "契"),
      }),
      check,
    ),
    {
      seed: 197,
      numRuns: 60,
      id: "replacement plans preserve range metadata across grapheme endpoints and ownership",
    },
  );
});
