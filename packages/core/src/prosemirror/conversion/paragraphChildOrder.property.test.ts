/**
 * An editor save writes a paragraph's children back in the order the source
 * held them.
 *
 * `toProseDoc` flattens a paragraph into one inline sequence: runs become
 * leaves, and `w:hyperlink`, `w:sdt`, `w:smartTag`, `w:fldSimple`, `w:ins`,
 * `w:del` and the bookmark markers become marks, atoms or boundary nodes on
 * it. `fromProseDoc` rebuilds the containers from that sequence, and every
 * container it opens has to be closed before the next sibling is written, or a
 * sibling that follows a container in the source lands ahead of it and the
 * visible text reads in a different order.
 *
 * Runs carry a revision-session attribute, as authored runs usually do, so
 * each one keeps its own identity through the editor instead of joining its
 * neighbour.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import type { Paragraph } from "../../types/document";
import { parseParagraph } from "../../docx/paragraphParser";
import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import { getLocalName, parseXmlDocument, type XmlElement } from "../../docx/xmlParser";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

setDefaultTimeout(propertyTestTimeout(30_000));

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REVISION = 'w:author="Reviewer" w:date="2026-09-25T10:00:00Z"';

type ChildKind =
  | "run"
  | "tabRun"
  | "hyperlink"
  | "linkWithSessionRuns"
  | "sdt"
  | "smartTag"
  | "fldSimple"
  | "ins"
  | "del"
  | "bookmarked";

const CHILD_KINDS = [
  "run",
  "tabRun",
  "hyperlink",
  "linkWithSessionRuns",
  "sdt",
  "smartTag",
  "fldSimple",
  "ins",
  "del",
  "bookmarked",
] as const satisfies readonly ChildKind[];

const rsid = (index: number): string => `00A${String(index).padStart(5, "0")}`;

const sessionRun = (text: string, index: number): string =>
  `<w:r w:rsidR="${rsid(index)}"><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** One paragraph child; `index` keeps every id, name and text distinct. */
const childXml = (kind: ChildKind, index: number): string => {
  const text = `${kind}${index} `;
  switch (kind) {
    case "run":
      return sessionRun(text, index);
    case "tabRun":
      return `<w:r><w:tab/><w:t xml:space="preserve">${text}</w:t></w:r>`;
    case "hyperlink":
      return `<w:hyperlink w:anchor="target${index}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
    case "linkWithSessionRuns":
      return (
        `<w:hyperlink w:anchor="target${index}">` +
        `${sessionRun(`${text}a `, index)}${sessionRun(`${text}b `, index + 1)}</w:hyperlink>`
      );
    case "sdt":
      return (
        `<w:sdt><w:sdtPr><w:id w:val="${index + 1}"/></w:sdtPr>` +
        `<w:sdtContent><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:sdtContent></w:sdt>`
      );
    case "smartTag":
      return (
        `<w:smartTag w:uri="urn:folio:smart" w:element="Party${index}">` +
        `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:smartTag>`
      );
    case "fldSimple":
      return (
        `<w:fldSimple w:instr=" DOCPROPERTY Title${index} ">` +
        `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:fldSimple>`
      );
    case "ins":
      return `<w:ins w:id="${index + 1}" ${REVISION}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
    case "del":
      return `<w:del w:id="${index + 1}" ${REVISION}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
    case "bookmarked":
      return (
        `<w:bookmarkStart w:id="${index + 1}" w:name="mark${index}"/>` +
        `${sessionRun(text, index)}<w:bookmarkEnd w:id="${index + 1}"/>`
      );
  }
};

/** Ids and texts step by two so a two-run hyperlink's second run stays unique. */
const paragraphXml = (kinds: readonly ChildKind[]): string =>
  `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">` +
  kinds.map((kind, position) => childXml(kind, position * 2)).join("") +
  "</w:p>";

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

const textOf = (element: XmlElement): string => {
  const name = getLocalName(element.name);
  if (name === "t" || name === "delText") {
    return (element.elements ?? []).map((child) => String(child.text ?? "")).join("");
  }
  return (element.elements ?? []).map(textOf).join("");
};

/**
 * The paragraph's children as `element:text`, adjacent runs folded into one
 * entry: the save may split or join runs, but it may not move a sibling across
 * one.
 */
const childOrderOf = (xml: string): string[] => {
  const root = parseXmlDocument(xml);
  if (root === null) {
    throw new Error("The saved paragraph is not well-formed");
  }
  const order: string[] = [];
  let pendingRunText: string | undefined;
  for (const child of root.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }
    const name = getLocalName(child.name);
    if (name === "pPr") {
      continue;
    }
    if (name === "r") {
      pendingRunText = (pendingRunText ?? "") + textOf(child);
      continue;
    }
    if (pendingRunText !== undefined) {
      order.push(`r:${pendingRunText}`);
      pendingRunText = undefined;
    }
    order.push(`${name}:${textOf(child)}`);
  }
  if (pendingRunText !== undefined) {
    order.push(`r:${pendingRunText}`);
  }
  return order;
};

const expectOrderPreserved = (kinds: readonly ChildKind[]): void => {
  const authored = paragraphXml(kinds);
  const expected = childOrderOf(serializeParagraph(parseParagraphXml(authored)));
  const saved = serializeParagraph(throughEditor(parseParagraphXml(authored)));
  expect(childOrderOf(saved)).toEqual(expected);
};

describe("paragraph child order through an editor save", () => {
  test("a run between two hyperlinks stays between them", () => {
    expectOrderPreserved(["run", "hyperlink", "run", "hyperlink", "run"]);
  });

  test("every pair of child kinds keeps its order", () => {
    for (const first of CHILD_KINDS) {
      for (const second of CHILD_KINDS) {
        expectOrderPreserved([first, second, "run"]);
        expectOrderPreserved(["run", first, second]);
      }
    }
  });

  test("mixed sequences of runs and inline containers keep their order", () => {
    fc.assert(
      fc.property(
        fc.array(fc.constantFrom(...CHILD_KINDS), { minLength: 1, maxLength: 8 }),
        (kinds) => {
          expectOrderPreserved(kinds);
        },
      ),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
