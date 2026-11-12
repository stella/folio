/**
 * Blanks next to a line break inside a paragraph are text, and Markdown keeps
 * them: CommonMark strips the blanks that start a line and takes the ones
 * before a line ending as the break's syntax, so the writer encodes them, and
 * reading the Markdown back gives the paragraph's text with every blank.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { Document, Paragraph, Run } from "../types/document";
import { fromMarkdown } from "./fromMarkdown";
import { toMarkdown } from "./index";

setDefaultTimeout(propertyTestTimeout(30_000));

const CLEAN = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  hyperlinks: "inline",
  footnotes: "strip",
} as const;

type Piece = { kind: "text"; text: string; bold: boolean } | { kind: "break" };

const pieceArbitrary: fc.Arbitrary<Piece> = fc.oneof(
  fc.record({
    kind: fc.constant("text" as const),
    text: fc
      .array(fc.constantFrom("a", "Z", " ", "  ", "\t"), {
        minLength: 1,
        maxLength: 4,
      })
      .map((parts) => parts.join("")),
    bold: fc.boolean(),
  }),
  fc.constant({ kind: "break" as const }),
);

const runOf = (piece: Piece): Run =>
  piece.kind === "break"
    ? { type: "run", content: [{ type: "break" }] }
    : {
        type: "run",
        content: [{ type: "text", text: piece.text }],
        ...(piece.bold && { formatting: { bold: true } }),
      };

/** Framed by words, so the paragraph's own edges hold no blank. */
const documentOf = (pieces: readonly Piece[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            runOf({ kind: "text", text: "Start", bold: false }),
            ...pieces.map(runOf),
            runOf({ kind: "text", text: "end", bold: false }),
          ],
        },
      ],
    },
  },
});

const textOf = (paragraph: Paragraph): string =>
  paragraph.content
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .map((content) => {
      if (content.type === "text") return content.text;
      if (content.type === "break") return "\n";
      return "";
    })
    .join("");

describe("blanks beside a line break", () => {
  test("read back from Markdown as written", async () => {
    await assertProperty(
      fc.property(
        fc.array(pieceArbitrary, { minLength: 1, maxLength: 6 }),
        (pieces) => {
          const source = documentOf(pieces);
          const markdown = toMarkdown(source, CLEAN);
          const [paragraph] = fromMarkdown(markdown).package.document.content;
          const [written] = source.package.document.content;
          if (paragraph?.type !== "paragraph" || written?.type !== "paragraph")
            throw new Error(`no paragraph read from ${markdown}`);
          expect(textOf(paragraph)).toBe(textOf(written));
        },
      ),
      { id: "blanks beside a line break survive Markdown", numRuns: 300 },
    );
  });
});
