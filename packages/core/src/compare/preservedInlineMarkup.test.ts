/**
 * Markup a package writes BETWEEN runs is not a run.
 *
 * `w:proofErr` and its neighbours are `w:p` children folio does not model and
 * carries verbatim. They serialize outside `w:r`, so they hold no `w:rPr` and
 * nothing about them is authored run formatting. Counting one as a
 * run-formatting carrier asks a comparison to line the base document's
 * proofing annotations up with the revised document's own, which no redline
 * can do: the round-trip check then refused a redline whose content was right.
 */

import { expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import type { BlockContent, ParagraphContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2024-03-01T00:00:00.000Z",
  onUnverified: "emit",
} as const;

const run = (text: string) =>
  ({ type: "run", content: [{ type: "text", text }] }) as const satisfies ParagraphContent;

/** One grammar-check annotation, exactly as Word writes it beside a run. */
const proofErr = (kind: "gramStart" | "gramEnd") =>
  ({
    type: "preservedInline",
    xml: `<w:proofErr w:type="${kind}"/>`,
    text: "",
  }) as const satisfies ParagraphContent;

const cover = (paraId: string, terms: string): BlockContent => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: [
    run("By signing this Order Form, each party agrees to "),
    proofErr("gramStart"),
    run("enter into"),
    proofErr("gramEnd"),
    run(` the ${terms} Terms.`),
  ],
});

const buildDocx = (terms: string) => {
  const result = createEmptyDocument();
  result.package.document.content = [
    cover("11111111", terms),
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            { type: "tableCell", content: [cover("22222222", terms)] },
            { type: "tableCell", content: [run("Payment Process")].map(asParagraph) },
          ],
        },
      ],
    },
  ];
  return createDocx(result);
};

const asParagraph = (content: ParagraphContent): BlockContent => ({
  type: "paragraph",
  paraId: "33333333",
  textId: "33333333",
  content: [content],
});

test("a paragraph carrying proofing markup round-trips through a word replacement", async () => {
  const result = await compareDocx(await buildDocx("Framework"), await buildDocx("Standard"), {
    ...OPTIONS,
    granularity: "word",
  });
  if (result.isErr()) {
    throw result.error;
  }

  expect(result.value.verification).toEqual({ status: "verified" });
  expect(result.value.changes).toHaveLength(2);
});
