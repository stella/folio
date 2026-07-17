/**
 * Class-guard tests for the suggestion serialization strip.
 *
 * A *suggested* tracked change (provenance "suggested") must NEVER reach
 * serialized DOCX output: `extractBlocks` strips it at the single PM->model
 * boundary that every serialization path funnels through. These tests build
 * documents that mix suggested changes with real (user) tracked changes and
 * assert the serialized model / XML contains no trace of the suggested content
 * while the real tracked changes survive unchanged.
 */

import { describe, expect, test } from "bun:test";
import type { Mark, Node as PMNode } from "prosemirror-model";

import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import type { Document, Paragraph } from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";

const DATE = "2026-07-17T00:00:00.000Z";

const insertion = (attrs: Record<string, unknown>): Mark =>
  // SAFETY: the editor schema always defines the insertion mark.
  schema.marks["insertion"]!.create({ author: "AI", date: DATE, ...attrs });

const deletion = (attrs: Record<string, unknown>): Mark =>
  // SAFETY: the editor schema always defines the deletion mark.
  schema.marks["deletion"]!.create({ author: "AI", date: DATE, ...attrs });

const paragraphOf = (nodes: PMNode[]): PMNode =>
  // SAFETY: the editor schema always defines paragraph + doc nodes.
  schema.nodes["doc"]!.create({}, [schema.nodes["paragraph"]!.create({}, nodes)]);

const firstParagraph = (doc: Document): Paragraph => {
  const block = doc.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected a paragraph block");
  }
  return block;
};

const paragraphPlainText = (paragraph: Paragraph): string => {
  const collect = (content: Paragraph["content"]): string =>
    content
      .map((item) => {
        if (item.type === "run") {
          return item.content.map((c) => (c.type === "text" ? c.text : "")).join("");
        }
        if (
          item.type === "insertion" ||
          item.type === "deletion" ||
          item.type === "moveTo" ||
          item.type === "moveFrom"
        ) {
          return collect(item.content);
        }
        if (item.type === "hyperlink") {
          return collect(item.children);
        }
        return "";
      })
      .join("");
  return collect(paragraph.content);
};

describe("suggestion serialization strip", () => {
  test("a suggested insertion is omitted entirely from the model and XML", () => {
    const doc = paragraphOf([
      schema.text("Keep this "),
      schema.text("suggested add", [
        insertion({ revisionId: 900, provenance: "suggested", suggestionId: "s1" }),
      ]),
      schema.text(" tail"),
    ]);

    const paragraph = firstParagraph(fromProseDoc(doc));
    expect(paragraphPlainText(paragraph)).toBe("Keep this  tail");
    expect(paragraph.content.some((item) => item.type === "insertion")).toBe(false);

    const xml = serializeParagraph(paragraph);
    expect(xml).not.toContain("suggested add");
    expect(xml).not.toContain("<w:ins");
  });

  test("a suggested deletion survives as plain content (removal never happened)", () => {
    const doc = paragraphOf([
      schema.text("Keep "),
      schema.text("proposed removal", [
        deletion({ revisionId: 901, provenance: "suggested", suggestionId: "s2" }),
      ]),
      schema.text(" here"),
    ]);

    const paragraph = firstParagraph(fromProseDoc(doc));
    expect(paragraphPlainText(paragraph)).toBe("Keep proposed removal here");
    expect(paragraph.content.some((item) => item.type === "deletion")).toBe(false);

    const xml = serializeParagraph(paragraph);
    expect(xml).toContain("proposed removal");
    expect(xml).not.toContain("<w:del");
  });

  test("real tracked changes adjacent to suggested ones serialize unchanged", () => {
    const doc = paragraphOf([
      schema.text("real ins", [insertion({ revisionId: 10 })]),
      schema.text("suggested ins", [
        insertion({ revisionId: 11, provenance: "suggested", suggestionId: "s3" }),
      ]),
      schema.text("real del", [deletion({ revisionId: 12 })]),
      schema.text("suggested del", [
        deletion({ revisionId: 13, provenance: "suggested", suggestionId: "s3" }),
      ]),
    ]);

    const paragraph = firstParagraph(fromProseDoc(doc));
    const xml = serializeParagraph(paragraph);

    // Real tracked changes survive.
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("real ins");
    expect(xml).toContain("<w:del");
    expect(xml).toContain("real del");

    // Suggested insertion text is gone; suggested deletion text is plain.
    expect(xml).not.toContain("suggested ins");
    expect(xml).toContain("suggested del");

    // Exactly one w:ins and one w:del (the real ones), never the suggested pair.
    expect(xml.match(/<w:ins\b/g)?.length ?? 0).toBe(1);
    expect(xml.match(/<w:del\b/g)?.length ?? 0).toBe(1);
  });

  test("suggested marks never leak even when they carry no suggestionId", () => {
    const doc = paragraphOf([
      schema.text("orphan suggested", [insertion({ revisionId: 950, provenance: "suggested" })]),
      schema.text(" survivor"),
    ]);

    const paragraph = firstParagraph(fromProseDoc(doc));
    expect(paragraphPlainText(paragraph)).toBe(" survivor");
    expect(serializeParagraph(paragraph)).not.toContain("orphan suggested");
  });
});
