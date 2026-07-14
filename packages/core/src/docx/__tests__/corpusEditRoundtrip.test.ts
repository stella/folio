/**
 * Every corpus fixture must survive a real document mutation, save, and reopen.
 * The inserted paragraph proves the model changed; byte comparisons prove that
 * package parts unrelated to the body edit were preserved.
 */

import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

const FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
const REGRESSION_FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "regressions");
const EDIT_MARKER_PREFIX = "Folio corpus edit: ";
const BODY_PART = "word/document.xml";

const FIXTURE_FILES = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

type ReadFixtureOptions = {
  directory?: string;
  filename: string;
};

const readFixture = ({ directory = FIXTURES_DIR, filename }: ReadFixtureOptions): ArrayBuffer => {
  const bytes = readFileSync(path.join(directory, filename));
  return new Uint8Array(bytes).buffer;
};

const packageFileNames = (zip: JSZip): string[] =>
  Object.values(zip.files)
    .filter((file) => !file.dir)
    .map((file) => file.name)
    .sort();

type ExpectUnrelatedPartsPreservedOptions = {
  originalBytes: ArrayBuffer;
  savedBytes: ArrayBuffer;
};

const expectUnrelatedPartsPreserved = async ({
  originalBytes,
  savedBytes,
}: ExpectUnrelatedPartsPreservedOptions): Promise<void> => {
  const [originalZip, savedZip] = await Promise.all([
    JSZip.loadAsync(originalBytes),
    JSZip.loadAsync(savedBytes),
  ]);
  const originalNames = packageFileNames(originalZip);
  expect(packageFileNames(savedZip)).toEqual(originalNames);

  for (const filename of originalNames) {
    if (filename === BODY_PART) {
      continue;
    }

    const originalFile = originalZip.file(filename);
    const savedFile = savedZip.file(filename);
    if (!originalFile || !savedFile) {
      throw new Error(`Expected package part is missing: ${filename}`);
    }

    const [originalPart, savedPart] = await Promise.all([
      originalFile.async("nodebuffer"),
      savedFile.async("nodebuffer"),
    ]);
    expect(savedPart.equals(originalPart)).toBe(true);
  }
};

type TextMatch = {
  from: number;
  node: ProseMirrorNode;
  parent: ProseMirrorNode | null;
};

const findUniqueText = (doc: ProseMirrorNode, text: string): TextMatch => {
  const matches: TextMatch[] = [];

  doc.descendants((node, from, parent) => {
    if (node.isText && node.text === text) {
      matches.push({ from, node, parent });
    }
  });

  if (matches.length !== 1) {
    panic(`Expected one text node matching "${text}", found ${matches.length}`);
  }

  return matches[0];
};

type ReplaceUniqueTextOptions = {
  doc: ProseMirrorNode;
  text: string;
  replacement: string;
};

const replaceUniqueText = ({
  doc,
  text,
  replacement,
}: ReplaceUniqueTextOptions): ProseMirrorNode => {
  const match = findUniqueText(doc, text);
  const replacementNode = doc.type.schema.text(replacement, match.node.marks);
  const state = EditorState.create({ doc });
  return state.apply(
    state.tr.replaceWith(match.from, match.from + match.node.nodeSize, replacementNode),
  ).doc;
};

type EditFixtureTextOptions = {
  directory?: string;
  filename: string;
  text: string;
  replacement: string;
};

const editFixtureText = async ({
  directory,
  filename,
  text,
  replacement,
}: EditFixtureTextOptions) => {
  const originalBytes = readFixture({ directory, filename });
  const parsed = await parseDocx(originalBytes);
  const originalPm = toProseDoc(parsed);
  const editedPm = replaceUniqueText({
    doc: originalPm,
    text,
    replacement,
  });
  const savedBytes = await repackDocx(fromProseDoc(editedPm, parsed), {
    updateModifiedDate: false,
  });
  const reopened = await parseDocx(savedBytes);

  return {
    originalBytes,
    originalPm,
    originalSectionProperties: parsed.package.document.sections.map(({ properties }) => properties),
    savedBytes,
    reopenedPm: toProseDoc(reopened),
    reopenedSectionProperties: reopened.package.document.sections.map(
      ({ properties }) => properties,
    ),
  };
};

const nodesOfType = (doc: ProseMirrorNode, type: string): ProseMirrorNode[] => {
  const matches: ProseMirrorNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === type) {
      matches.push(node);
    }
  });
  return matches;
};

const sectionBreakProperties = (doc: ProseMirrorNode): unknown[] => {
  const properties: unknown[] = [];
  doc.forEach((node) => {
    const sectionProperties = node.attrs["_sectionProperties"];
    if (sectionProperties !== null && sectionProperties !== undefined) {
      properties.push(sectionProperties);
    }
  });
  return properties;
};

describe("corpus edit/save/reopen", () => {
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)("preserves package content after a body edit (%s)", async (filename) => {
    const originalBytes = readFixture({ filename });
    const parsed = await parseDocx(originalBytes);
    const originalPm = toProseDoc(parsed);
    const marker = `${EDIT_MARKER_PREFIX}${filename}`;
    const markerParagraph = originalPm.type.schema.nodes.paragraph.create(
      null,
      originalPm.type.schema.text(marker),
    );
    const state = EditorState.create({ doc: originalPm });
    const editedPm = state.apply(state.tr.insert(state.doc.content.size, markerParagraph)).doc;
    const savedBytes = await repackDocx(fromProseDoc(editedPm, parsed), {
      updateModifiedDate: false,
    });
    const reopenedPm = toProseDoc(await parseDocx(savedBytes));

    expect(reopenedPm.textContent).toBe(`${originalPm.textContent}${marker}`);
    expect(reopenedPm.childCount).toBe(originalPm.childCount + 1);
    await expectUnrelatedPartsPreserved({ originalBytes, savedBytes });
  });
});

describe("structured corpus edits", () => {
  test("edits text inside a table cell", async () => {
    const result = await editFixtureText({
      filename: "upstream-with-tables.docx",
      text: "B2",
      replacement: "B2 edited",
    });
    const tables = nodesOfType(result.reopenedPm, "table");

    expect(tables).toHaveLength(1);
    expect(tables[0].textContent).toContain("B2 edited");
    expect(nodesOfType(tables[0], "tableRow")).toHaveLength(3);
    expect(nodesOfType(tables[0], "tableCell")).toHaveLength(9);
    await expectUnrelatedPartsPreserved(result);
  });

  test("preserves bold formatting on edited text", async () => {
    const result = await editFixtureText({
      filename: "upstream-styled-content.docx",
      text: "Bold text. ",
      replacement: "Bold text edited. ",
    });
    const match = findUniqueText(result.reopenedPm, "Bold text edited. ");

    expect(match.node.marks.map((mark) => mark.type.name)).toContain("bold");
    await expectUnrelatedPartsPreserved(result);
  });

  test("preserves the paragraph style on edited heading text", async () => {
    const result = await editFixtureText({
      filename: "upstream-complex-styles.docx",
      text: "Heading 1",
      replacement: "Heading 1 edited",
    });
    const match = findUniqueText(result.reopenedPm, "Heading 1 edited");

    expect(match.parent?.attrs["styleId"]).toBe("Heading1");
    await expectUnrelatedPartsPreserved(result);
  });

  test("preserves a section boundary after editing adjacent text", async () => {
    const result = await editFixtureText({
      directory: REGRESSION_FIXTURES_DIR,
      filename: "repack-paragraph-sectpr.docx",
      text: "Možnosti odstoupení objednatele od objednávky:",
      replacement: "Možnosti odstoupení objednatele od objednávky: upraveno",
    });
    const originalSections = sectionBreakProperties(result.originalPm);

    expect(originalSections).toHaveLength(1);
    expect(sectionBreakProperties(result.reopenedPm)).toEqual(originalSections);
    expect(result.originalSectionProperties).toHaveLength(2);
    expect(result.reopenedSectionProperties).toEqual(result.originalSectionProperties);
    findUniqueText(result.reopenedPm, "Možnosti odstoupení objednatele od objednávky: upraveno");
    await expectUnrelatedPartsPreserved(result);
  });

  test("preserves drawings and section geometry after a text edit", async () => {
    const result = await editFixtureText({
      directory: REGRESSION_FIXTURES_DIR,
      filename: "repack-image-count.docx",
      text: "2026-PU-036",
      replacement: "2026-PU-036A",
    });
    const originalSections = sectionBreakProperties(result.originalPm);

    expect(nodesOfType(result.originalPm, "image")).toHaveLength(43);
    expect(nodesOfType(result.reopenedPm, "image")).toHaveLength(43);
    expect(originalSections).toHaveLength(7);
    expect(sectionBreakProperties(result.reopenedPm)).toEqual(originalSections);
    expect(result.originalSectionProperties).toHaveLength(8);
    expect(result.reopenedSectionProperties).toEqual(result.originalSectionProperties);
    findUniqueText(result.reopenedPm, "2026-PU-036A");
    await expectUnrelatedPartsPreserved(result);
  });
});
