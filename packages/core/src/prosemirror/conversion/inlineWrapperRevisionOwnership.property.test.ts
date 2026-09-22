/**
 * A transparent wrapper and a tracked revision each own their position in the
 * OOXML tree. Projecting both as marks must retain which wrapper layers were
 * outside the revision and which were inside it.
 */

import { describe, expect, test } from "bun:test";

import type { Paragraph } from "../../types/document";
import { parseParagraph } from "../../docx/paragraphParser";
import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import { parseXmlDocument } from "../../docx/xmlParser";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type WrapperFixture = {
  name: "bdo" | "dir" | "smartTag" | "customXml";
  wrap: (content: string) => string;
};

const WRAPPERS = [
  { name: "bdo", wrap: (content) => `<w:bdo w:val="rtl">${content}</w:bdo>` },
  { name: "dir", wrap: (content) => `<w:dir w:val="ltr">${content}</w:dir>` },
  {
    name: "smartTag",
    wrap: (content) =>
      `<w:smartTag w:uri="urn:folio:smart" w:element="Party">${content}</w:smartTag>`,
  },
  {
    name: "customXml",
    wrap: (content) =>
      `<w:customXml w:uri="urn:folio:custom" w:element="Matter">${content}</w:customXml>`,
  },
] as const satisfies readonly WrapperFixture[];

type RevisionFixture = {
  name: "ins" | "del" | "moveFrom" | "moveTo";
  id: number;
  textElement: "t" | "delText";
};

const REVISIONS = [
  { name: "ins", id: 101, textElement: "t" },
  { name: "del", id: 102, textElement: "delText" },
  { name: "moveFrom", id: 103, textElement: "delText" },
  { name: "moveTo", id: 104, textElement: "t" },
] as const satisfies readonly RevisionFixture[];

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml);
  if (root === null) {
    throw new Error("Failed to parse the paragraph fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
};

const throughEditor = (paragraph: Paragraph): Paragraph => {
  const source = { package: { document: { content: [paragraph] } } };
  const block = fromProseDoc(toProseDoc(source), source).package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The editor round trip lost its paragraph");
  }
  return block;
};

const saveThroughEditor = (xml: string): string =>
  serializeParagraph(throughEditor(parseParagraphXml(xml)));

const paragraph = (body: string): string =>
  `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">${body}</w:p>`;

const revisionXml = ({ id, name }: RevisionFixture, content: string): string =>
  `<w:${name} w:id="${id}" w:author="Reviewer ${id}" w:date="2026-09-22T10:00:00Z">` +
  `${content}</w:${name}>`;

const revisionRun = ({ id, textElement }: RevisionFixture): string =>
  `<w:r><w:${textElement}>revision ${id}</w:${textElement}></w:r>`;

const plainRun = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;

const expectExactFixedPoint = (authored: string): void => {
  const expected = serializeParagraph(parseParagraphXml(authored));
  const once = saveThroughEditor(authored);
  expect(once).toBe(expected);
  expect(saveThroughEditor(paragraph(once.slice(once.indexOf(">") + 1, -6)))).toBe(once);
};

describe("transparent-wrapper ownership around tracked revisions", () => {
  test("is exact and reaches a fixed point for every wrapper and revision kind", () => {
    const exercisedOuterWrappers = new Set<string>();
    const exercisedInnerWrappers = new Set<string>();
    const exercisedRevisions = new Set<string>();

    for (const [outerIndex, outer] of WRAPPERS.entries()) {
      const inner = WRAPPERS[(outerIndex + 1) % WRAPPERS.length];
      if (inner === undefined) {
        throw new Error("The wrapper matrix did not provide an inner wrapper");
      }
      exercisedOuterWrappers.add(outer.name);
      exercisedInnerWrappers.add(inner.name);

      for (const revision of REVISIONS) {
        exercisedRevisions.add(revision.name);
        const authored = paragraph(
          outer.wrap(revisionXml(revision, inner.wrap(revisionRun(revision)))),
        );
        expectExactFixedPoint(authored);
      }
    }

    const wrapperNames = WRAPPERS.map(({ name }) => name);
    expect([...exercisedOuterWrappers].toSorted()).toEqual([...wrapperNames].toSorted());
    expect([...exercisedInnerWrappers].toSorted()).toEqual([...wrapperNames].toSorted());
    expect([...exercisedRevisions].toSorted()).toEqual(
      REVISIONS.map(({ name }) => name).toSorted(),
    );
  });

  test("keeps one outer wrapper around plain runs and consecutive distinct revisions", () => {
    const insertion = REVISIONS.at(0);
    const deletion = REVISIONS.at(1);
    if (insertion === undefined || deletion === undefined) {
      throw new Error("The revision matrix did not provide two distinct revisions");
    }
    const authored = paragraph(
      WRAPPERS[0].wrap(
        plainRun("before") +
          revisionXml(insertion, revisionRun(insertion)) +
          revisionXml(deletion, revisionRun(deletion)) +
          plainRun("after"),
      ),
    );

    expectExactFixedPoint(authored);
  });

  test("keeps a revision outside each transparent wrapper kind", () => {
    for (const wrapper of WRAPPERS) {
      for (const revision of REVISIONS) {
        expectExactFixedPoint(
          paragraph(revisionXml(revision, wrapper.wrap(revisionRun(revision)))),
        );
      }
    }
  });
});
