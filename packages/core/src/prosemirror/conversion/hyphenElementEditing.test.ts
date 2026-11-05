/**
 * What editing does to the mark that keeps a hyphen element an element: typed
 * text does not take it, and a new field result does not keep the old one's.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import type {
  ComplexField,
  Document,
  Paragraph,
  Run,
  RunContent,
} from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import {
  HYPHEN_ELEMENT_MARK_NAME,
  HYPHEN_TEXT_CARRIERS,
} from "./hyphenTextCarriers";
import { toProseDoc } from "./toProseDoc";

const SAVE_OPTIONS = { stylesheetSource: { type: "package" } } as const;

const documentOf = (content: Paragraph["content"]): Document => ({
  package: { document: { content: [{ type: "paragraph", content }] } },
});

const savedParagraph = (doc: PMNode): Paragraph => {
  const paragraph = fromProseDoc(doc, undefined, SAVE_OPTIONS).package.document
    .content[0];
  return paragraph?.type === "paragraph"
    ? paragraph
    : panic("Expected a saved paragraph");
};

const savedRunContent = (doc: PMNode): RunContent[] =>
  savedParagraph(doc).content.flatMap((item) =>
    item.type === "run" ? item.content : [],
  );

const carriers: RunContent[] = [
  { type: "softHyphen" },
  { type: "softHyphen" },
  { type: "noBreakHyphen" },
];

/** Paragraph start is 1, so the first carrier ends at 2 and the second at 3. */
const BETWEEN_CARRIERS = 2;

describe("typing next to a hyphen element", () => {
  test("adjacent carriers are separate leaves, even of one kind", () => {
    const doc = toProseDoc(documentOf([{ type: "run", content: carriers }]));
    expect(doc.child(0).childCount).toBe(carriers.length);
  });

  test.each([
    ["a soft hyphen", HYPHEN_TEXT_CARRIERS.softHyphen],
    ["a non-breaking hyphen", HYPHEN_TEXT_CARRIERS.noBreakHyphen],
    ["a letter", "x"],
  ])("%s typed between two carriers stays text", (_label, typed) => {
    const doc = toProseDoc(documentOf([{ type: "run", content: carriers }]));
    const typedDoc = EditorState.create({ doc }).tr.insertText(
      typed,
      BETWEEN_CARRIERS,
    ).doc;

    expect(
      typedDoc
        .child(0)
        .child(1)
        .marks.map((mark) => mark.type.name),
    ).not.toContain(HYPHEN_ELEMENT_MARK_NAME);
    expect(savedRunContent(typedDoc)).toEqual([
      { type: "softHyphen" },
      { type: "text", text: typed },
      { type: "softHyphen" },
      { type: "noBreakHyphen" },
    ]);
  });

  test("text replacing a carrier stays text", () => {
    const doc = toProseDoc(documentOf([{ type: "run", content: carriers }]));
    const typedDoc = EditorState.create({ doc }).tr.insertText(
      HYPHEN_TEXT_CARRIERS.softHyphen,
      BETWEEN_CARRIERS,
      BETWEEN_CARRIERS + 1,
    ).doc;

    expect(savedRunContent(typedDoc)).toEqual([
      { type: "softHyphen" },
      { type: "text", text: HYPHEN_TEXT_CARRIERS.softHyphen },
      { type: "noBreakHyphen" },
    ]);
  });
});

const fieldWithResult = (fieldResult: Run[]): ComplexField => ({
  type: "complexField",
  fieldType: "REF",
  instruction: "REF target \\r",
  fieldCode: [],
  fieldResult,
});

const savedFieldResult = (
  field: ComplexField,
  displayText: string,
): Paragraph["content"] => {
  const doc = toProseDoc(documentOf([field]));
  const node = doc.child(0).child(0);
  expect(node.type.name).toBe("structuredField");
  const refreshed = schema.node("doc", null, [
    schema.node("paragraph", null, [
      node.type.create(
        { ...node.attrs, displayText },
        node.content,
        node.marks,
      ),
    ]),
  ]);
  const saved = savedParagraph(refreshed).content.at(0);
  return saved?.type === "complexField"
    ? saved.fieldResult
    : panic("Expected a saved complex field");
};

const resultContent = (result: Paragraph["content"]): RunContent[] =>
  result.flatMap((item) => (item.type === "run" ? item.content : []));

describe("a new field display text", () => {
  test("replaces the hyphen elements of the old result", () => {
    const result = savedFieldResult(
      fieldWithResult([
        {
          type: "run",
          content: [
            { type: "text", text: "Article" },
            { type: "noBreakHyphen" },
            { type: "softHyphen" },
            { type: "text", text: "1" },
          ],
        },
      ]),
      "2",
    );

    expect(resultContent(result).filter(({ type }) => type !== "text")).toEqual(
      [],
    );
    expect(
      resultContent(result)
        .map((item) => (item.type === "text" ? item.text : ""))
        .join(""),
    ).toBe("2");
  });

  test("takes the place of a result that held only hyphen elements", () => {
    const result = savedFieldResult(
      fieldWithResult([{ type: "run", content: [{ type: "noBreakHyphen" }] }]),
      "2",
    );

    expect(resultContent(result)).toEqual([{ type: "text", text: "2" }]);
  });

  test("leaves an unchanged result's hyphen elements alone", () => {
    const content: RunContent[] = [
      { type: "text", text: "Article" },
      { type: "noBreakHyphen" },
      { type: "text", text: "1" },
    ];
    const result = savedFieldResult(
      fieldWithResult([{ type: "run", content }]),
      `Article${HYPHEN_TEXT_CARRIERS.noBreakHyphen}1`,
    );

    expect(resultContent(result)).toEqual(content);
  });
});
