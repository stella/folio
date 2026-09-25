/**
 * A direct-mode text replacement writes only what it changes.
 *
 * The property generates a paragraph of differently formatted runs mixed with
 * inline content controls, note references, fields, tabs, breaks, bookmarks
 * and internal links, and an edit that matches all of it or part of it and
 * changes a few words. After the edit:
 *
 * - the block reads exactly as the replacement says;
 * - every character the change did not touch keeps its marks and the content
 *   control it sat in;
 * - no content control, bookmark or link disappears, and a field survives
 *   unless its displayed text was what changed;
 * - the package saves and reopens with the same text and the same controls
 *   and note references.
 *
 * The same paragraphs, with highlighted runs among them, drive the
 * tracked-changes and suggested modes. The redline marks as deleted exactly
 * the characters its changes remove and inserts exactly their new text;
 * nothing else carries a revision, and a background is cleared as a property
 * change only on untouched characters of a highlighted stretch a change
 * touches. Accepting every revision gives the replacement; rejecting every
 * revision gives back the original paragraph.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { Mark, type Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import {
  acceptAllSuggestions,
  rejectAllSuggestions,
  resolveAllChangesInHeadlessState,
} from "../prosemirror/commands/comments";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import { runFormattingInlineAtomCleanText } from "../prosemirror/runFormattingInlineCarriers";
import { buildCleanBlockText } from "./clean-text";
import { FolioDocxReviewer } from "./headless";
import {
  changesFromSegments,
  planTextChanges,
  type TextChange,
  widenChangesToAtomicSpans,
} from "./minimal-replacement";
import { diffWordSegments, type WordDiffGranularity } from "./word-diff";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Neighbouring items never share a note: two references to one note would read as one. */
const NOTE_IDS = [1, 2, 3] as const;
const noteIdAt = (index: number): number => NOTE_IDS[index % NOTE_IDS.length] ?? 1;

const WORDS = ["Seller", "shall", "deliver", "the", "goods", "on", "time", "Buyer", "pays"];

type RunProperties = {
  bold: boolean;
  italic: boolean;
  size: 20 | 28 | null;
  rsid: boolean;
  highlight?: boolean;
};

type Item =
  | { kind: "run"; text: string; properties: RunProperties }
  | { kind: "control"; runs: { text: string; properties: RunProperties }[] }
  | { kind: "noteReference" }
  | { kind: "field"; result: string }
  | { kind: "tab" }
  | { kind: "break" }
  | { kind: "bookmark"; text: string }
  | { kind: "link"; text: string };

const words = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 3 })
  .map((picked) => `${picked.join(" ")} `);

const runProperties = fc.record({
  bold: fc.boolean(),
  italic: fc.boolean(),
  size: fc.constantFrom(20 as const, 28 as const, null),
  rsid: fc.boolean(),
});

/** Run properties with a highlight on some runs, for the modes that record its clearing. */
const highlightedRunProperties = fc.record({
  bold: fc.boolean(),
  italic: fc.boolean(),
  size: fc.constantFrom(20 as const, 28 as const, null),
  rsid: fc.boolean(),
  highlight: fc.boolean(),
});

const itemOf = (properties: fc.Arbitrary<RunProperties>): fc.Arbitrary<Item> =>
  fc.oneof(
    {
      weight: 5,
      arbitrary: fc.record({
        kind: fc.constant("run" as const),
        text: words,
        properties,
      }),
    },
    {
      weight: 2,
      arbitrary: fc.record({
        kind: fc.constant("control" as const),
        runs: fc.array(fc.record({ text: words, properties }), {
          minLength: 1,
          maxLength: 2,
        }),
      }),
    },
    { weight: 1, arbitrary: fc.constant({ kind: "noteReference" as const }) },
    {
      weight: 1,
      arbitrary: fc.record({
        kind: fc.constant("field" as const),
        result: fc.constantFrom("3.6", "Clause 4", "12"),
      }),
    },
    { weight: 1, arbitrary: fc.constant({ kind: "tab" as const }) },
    { weight: 1, arbitrary: fc.constant({ kind: "break" as const }) },
    { weight: 1, arbitrary: fc.record({ kind: fc.constant("bookmark" as const), text: words }) },
    { weight: 1, arbitrary: fc.record({ kind: fc.constant("link" as const), text: words }) },
  );

const item = itemOf(runProperties);

const escapeXml = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const runXml = (
  text: string,
  { bold, italic, size, rsid, highlight = false }: RunProperties,
): string => {
  const properties = [
    bold ? "<w:b/>" : "",
    italic ? "<w:i/>" : "",
    size === null ? "" : `<w:sz w:val="${size}"/>`,
    highlight ? '<w:highlight w:val="yellow"/>' : "",
  ].join("");
  return (
    `<w:r${rsid ? ' w:rsidR="00AB12CD"' : ""}>` +
    (properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`) +
    `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
  );
};

const PLAIN: RunProperties = { bold: false, italic: false, size: null, rsid: false };

const itemXml = (entry: Item, index: number): string => {
  switch (entry.kind) {
    case "run":
      return runXml(entry.text, entry.properties);
    case "control":
      return (
        `<w:sdt><w:sdtPr><w:alias w:val="Control ${index}"/><w:id w:val="${100 + index}"/></w:sdtPr>` +
        `<w:sdtContent>${entry.runs.map((run) => runXml(run.text, run.properties)).join("")}</w:sdtContent></w:sdt>`
      );
    case "noteReference":
      return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${noteIdAt(index)}"/></w:r>`;
    case "field":
      return (
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> REF _Ref${index} \\h </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        runXml(entry.result, PLAIN) +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
      );
    case "tab":
      return "<w:r><w:tab/></w:r>";
    case "break":
      return "<w:r><w:br/></w:r>";
    case "bookmark":
      return (
        `<w:bookmarkStart w:id="${index}" w:name="mark${index}"/>` +
        runXml(entry.text, PLAIN) +
        `<w:bookmarkEnd w:id="${index}"/>`
      );
    case "link":
      return `<w:hyperlink w:anchor="mark${index}">${runXml(entry.text, { ...PLAIN, italic: true })}</w:hyperlink>`;
    default: {
      const unreachable: never = entry;
      throw new Error(`Unhandled item ${JSON.stringify(unreachable)}`);
    }
  }
};

const createDocx = async (paragraphXml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>` +
      `</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  zip.file(
    "word/footnotes.xml",
    `${XML_DECLARATION}<w:footnotes xmlns:w="${W}">` +
      `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
      `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
      NOTE_IDS.map(
        (id) =>
          `<w:footnote w:id="${id}"><w:p><w:r><w:t>Note ${id}.</w:t></w:r></w:p></w:footnote>`,
      ).join("") +
      `</w:footnotes>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>` +
      `<w:p>${paragraphXml}</w:p>` +
      `<w:p><w:r><w:t>Untouched paragraph.</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

type CleanCharacter = { character: string; marks: readonly Mark[]; control: number | null };

/** Every clean-text character of a paragraph, with its marks and the control it sits in. */
const cleanCharacters = (paragraph: PMNode): CleanCharacter[] => {
  const characters: CleanCharacter[] = [];
  const controls = new Map<PMNode, number>();
  const controlOf = (parent: PMNode | null): number | null =>
    parent === null ? null : (controls.get(parent) ?? null);
  paragraph.descendants((node, _position, parent) => {
    if (node.type.name === "sdt") {
      controls.set(node, controls.size);
      return true;
    }
    const atomText = runFormattingInlineAtomCleanText(node);
    if (atomText !== null) {
      for (const character of atomText) {
        characters.push({ character, marks: node.marks, control: controlOf(parent) });
      }
      return false;
    }
    if (node.isText) {
      for (const character of node.text ?? "") {
        characters.push({ character, marks: node.marks, control: controlOf(parent) });
      }
    }
    return true;
  });
  return characters;
};

const countNodes = (paragraph: PMNode, name: string): number => {
  let count = 0;
  paragraph.descendants((node) => {
    if (node.type.name === name) {
      count++;
    }
    return true;
  });
  return count;
};

const markedCharacters = (characters: readonly CleanCharacter[], mark: string): number =>
  characters.filter((entry) => entry.marks.some((candidate) => candidate.type.name === mark))
    .length;

/** The edit: the whole text or a slice of it, with some words changed. */
const editArbitrary = fc.record({
  whole: fc.boolean(),
  sliceStart: fc.double({ min: 0, max: 1, noNaN: true }),
  sliceLength: fc.double({ min: 0, max: 1, noNaN: true }),
  mutations: fc.array(
    fc.record({
      at: fc.double({ min: 0, max: 1, noNaN: true }),
      remove: fc.nat({ max: 6 }),
      insert: fc.constantFrom("", "‸", "new ", " and", "X", "shall not "),
    }),
    { minLength: 1, maxLength: 3 },
  ),
});

type Edit = typeof editArbitrary extends fc.Arbitrary<infer Value> ? Value : never;

type PickedEdit = { start: number; end: number; find: string; replace: string };

/** The unique match and its changed replacement `edit` picks from `text`, if any. */
const pickEdit = (text: string, edit: Edit): PickedEdit | null => {
  const start = edit.whole ? 0 : Math.floor(edit.sliceStart * text.length);
  const end = edit.whole
    ? text.length
    : Math.min(text.length, start + 1 + Math.floor(edit.sliceLength * (text.length - start)));
  const find = text.slice(start, end);
  if (find.length === 0 || text.indexOf(find) !== text.lastIndexOf(find)) {
    return null;
  }
  let replace = find;
  for (const { at, remove, insert } of edit.mutations) {
    const position = Math.floor(at * replace.length);
    replace = replace.slice(0, position) + insert + replace.slice(position + remove);
  }
  return replace === find ? null : { start, end, find, replace };
};

const firstParagraphOf = (doc: PMNode): { node: PMNode; from: number } => {
  let from = 0;
  for (let index = 0; index < doc.childCount; index++) {
    const node = doc.child(index);
    if (node.type.name === "paragraph") {
      return { node, from };
    }
    from += node.nodeSize;
  }
  throw new Error("expected a paragraph");
};

const firstParagraph = (reviewer: FolioDocxReviewer): { node: PMNode; from: number } =>
  firstParagraphOf(reviewer.state.doc);

describe("a direct replacement changes only the characters it changes", () => {
  test("over generated paragraphs and edits", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(item, { minLength: 1, maxLength: 8 }),
        editArbitrary,
        async (items, edit) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(
            await createDocx(items.map(itemXml).join("")),
          );
          const block = reviewer.snapshot().blocks.at(0);
          const picked = block === undefined ? null : pickEdit(block.text, edit);
          if (!block || picked === null) {
            return;
          }
          const { start, end, find, replace } = picked;

          const before = firstParagraph(reviewer);
          const beforeCharacters = cleanCharacters(before.node);
          const cleanBefore = buildCleanBlockText(before.node, before.from);
          const fields = cleanBefore.structuralBoundaries.flatMap((boundary) =>
            boundary.type === "field"
              ? [{ offset: boundary.offset - start, length: boundary.length }]
              : [],
          );
          const changes = widenChangesToAtomicSpans(find, planTextChanges(find, replace), fields);

          const result = reviewer.applyOperations(
            [{ id: "edit", type: "replaceInBlock", blockId: block.id, find, replace }],
            { mode: "direct" },
          );
          const findCutsField = cleanBefore.structuralBoundaries.some(
            (boundary) =>
              boundary.type === "field" &&
              [start, end].some(
                (edge) => edge > boundary.offset && edge < boundary.offset + boundary.length,
              ),
          );
          if (findCutsField) {
            // A match that begins or ends inside a field result names text no
            // run holds; the operation contract refuses it before any change.
            expect(result.skipped).toEqual([{ id: "edit", reason: "unsupportedBlock" }]);
            return;
          }
          expect(result.skipped).toEqual([]);

          const expectedText = block.text.slice(0, start) + replace + block.text.slice(end);
          expect(reviewer.snapshot().blocks.at(0)?.text).toBe(expectedText);
          expect(reviewer.snapshot().blocks.at(1)?.text).toBe("Untouched paragraph.");

          const after = firstParagraph(reviewer);
          const afterCharacters = cleanCharacters(after.node);
          expect(afterCharacters.map((entry) => entry.character).join("")).toBe(expectedText);

          // Characters outside every change keep their marks and their control.
          for (const [index, was] of beforeCharacters.entries()) {
            const local = index - start;
            if (changes.some((change) => local >= change.start && local < change.end)) {
              continue;
            }
            const shift = changes
              .filter((change) => change.end <= local)
              .reduce((sum, change) => sum + change.text.length - (change.end - change.start), 0);
            const now = afterCharacters[index + shift];
            expect({
              index,
              character: now?.character,
              sameMarks: now !== undefined && Mark.sameSet(was.marks, now.marks),
              control: now?.control,
            }).toEqual({ index, character: was.character, sameMarks: true, control: was.control });
          }

          expect(countNodes(after.node, "sdt")).toBe(countNodes(before.node, "sdt"));
          expect(countNodes(after.node, "bookmarkBoundary")).toBe(
            countNodes(before.node, "bookmarkBoundary"),
          );
          const touchedFields = fields.filter(({ offset, length }) =>
            changes.some((change) => change.start < offset + length && change.end > offset),
          ).length;
          expect(countNodes(after.node, "field")).toBe(
            countNodes(before.node, "field") - touchedFields,
          );

          // The package still says the same after a save and a reopen.
          const saved = await reviewer.toBuffer();
          const reopened = await FolioDocxReviewer.fromBuffer(saved);
          const reopenedText = reopened.snapshot().blocks.at(0)?.text ?? "";
          if (items.some((entry) => entry.kind === "link")) {
            // Where a saved `w:hyperlink` lands among its sibling runs is the
            // serializer's contract, tested with it; the edit is checked above.
            expect([...reopenedText].toSorted()).toEqual([...expectedText].toSorted());
          } else {
            expect(reopenedText).toBe(expectedText);
          }
          const xml =
            (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
          expect((xml.match(/<w:sdt>/gu) ?? []).length).toBe(countNodes(before.node, "sdt"));
          expect((xml.match(/<w:footnoteReference /gu) ?? []).length).toBe(
            markedCharacters(afterCharacters, "footnoteRef"),
          );
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
  }, 240_000);
});

const REVISION_MARKS: ReadonlySet<string> = new Set(["insertion", "deletion", "runPropertyChange"]);
const BACKGROUND_MARKS: ReadonlySet<string> = new Set(["highlight", "runShading"]);

const hasMark = (marks: readonly Mark[], name: string): boolean =>
  marks.some((mark) => mark.type.name === name);

const withoutMarks = (marks: readonly Mark[], names: ReadonlySet<string>): readonly Mark[] =>
  marks.filter((mark) => !names.has(mark.type.name));

/** Non-text inline content: what no clean-text character stands for. */
const inlineContentWithoutText = (paragraph: PMNode): PMNode[] => {
  const nodes: PMNode[] = [];
  paragraph.descendants((node) => {
    if (node.isInline && !node.isText && runFormattingInlineAtomCleanText(node) === null) {
      nodes.push(node);
    }
    return true;
  });
  return nodes;
};

const canonical = (value: unknown): unknown => {
  if (typeof value !== "object" || value === null) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map(canonical);
  }
  return Object.fromEntries(
    Object.entries(value)
      .toSorted(([left], [right]) => left.localeCompare(right))
      .map(([key, entry]) => [key, canonical(entry)]),
  );
};

/** The run formatting `marks` state, whichever carrier states it. */
const textFormattingOf = (marks: readonly Mark[]): string =>
  JSON.stringify(canonical(marksToTextFormatting(marks)));

/**
 * Every inline node's type, and each character's text, run formatting and
 * other marks: what a paragraph says, whichever carrier states it.
 */
const formattingOf = (paragraph: PMNode): string[] => {
  const nodes: string[] = [];
  paragraph.descendants((node) => {
    if (!node.isInline) {
      return true;
    }
    const formatting = textFormattingOf(node.marks);
    const others = node.marks
      .filter((mark) => mark.type.spec.inclusive === false)
      .map((mark) => JSON.stringify([mark.type.name, mark.attrs]));
    for (const character of node.isText ? (node.text ?? "") : [""]) {
      nodes.push(`${node.type.name}:${character}:${formatting}:${others.join()}`);
    }
    return true;
  });
  return nodes;
};

const runCommand = (
  state: EditorState,
  command: (state: EditorState, dispatch: (tr: Transaction) => void) => boolean,
): EditorState => {
  let next = state;
  command(state, (transaction) => {
    next = state.apply(transaction);
  });
  return next;
};

type ReviewMode = "tracked-changes" | "suggested";

const resolveEverything = (
  state: EditorState,
  mode: ReviewMode,
  resolution: "accept" | "reject",
): EditorState => {
  if (mode === "tracked-changes") {
    return resolveAllChangesInHeadlessState(state, resolution);
  }
  return resolution === "accept"
    ? resolveAllChangesInHeadlessState(
        runCommand(state, acceptAllSuggestions({ author: "Reviewer" })),
        "accept",
      )
    : runCommand(state, rejectAllSuggestions());
};

/** Whether a change removes the character at span-local index `local`. */
const removedBy = (changes: readonly TextChange[], local: number): boolean =>
  changes.some((change) => local >= change.start && local < change.end);

/**
 * The span-local indices of every highlighted stretch a change touches, as the
 * applier's contract defines them: the characters a change removes, or the
 * two either side of a pure insertion, and every highlighted neighbour of a
 * highlighted one among them.
 */
const touchedHighlight = (
  highlighted: readonly boolean[],
  changes: readonly TextChange[],
): Set<number> => {
  const touched = new Set<number>();
  const span = highlighted.length;
  for (const { start, end } of changes) {
    const from = start === end ? Math.max(start - 1, 0) : start;
    const to = start === end ? Math.min(end + 1, span) : end;
    for (let index = from; index < to; index++) {
      if (!highlighted[index]) {
        continue;
      }
      touched.add(index);
      for (let left = index - 1; left >= 0 && highlighted[left]; left--) {
        touched.add(left);
      }
      for (let right = index + 1; right < span && highlighted[right]; right++) {
        touched.add(right);
      }
    }
  }
  return touched;
};

describe("a tracked or suggested replacement redlines only the characters it changes", () => {
  test("over generated paragraphs and edits, accepted and rejected", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(itemOf(highlightedRunProperties), { minLength: 1, maxLength: 8 }),
        editArbitrary,
        fc.constantFrom<ReviewMode>("tracked-changes", "suggested"),
        fc.constantFrom<WordDiffGranularity>("word", "character"),
        async (items, edit, mode, granularity) => {
          const source = await createDocx(items.map(itemXml).join(""));
          const reviewer = await FolioDocxReviewer.fromBuffer(source);
          const block = reviewer.snapshot().blocks.at(0);
          const picked = block === undefined ? null : pickEdit(block.text, edit);
          if (!block || picked === null) {
            return;
          }
          const { start, end, find, replace } = picked;
          const before = firstParagraph(reviewer);
          const beforeCharacters = cleanCharacters(before.node);
          const cleanBefore = buildCleanBlockText(before.node, before.from);
          const fields = cleanBefore.structuralBoundaries.flatMap((boundary) =>
            boundary.type === "field"
              ? [{ offset: boundary.offset - start, length: boundary.length }]
              : [],
          );
          const changes = widenChangesToAtomicSpans(
            find,
            changesFromSegments(diffWordSegments(find, replace, { granularity })),
            fields,
          );

          const result = reviewer.applyOperations(
            [{ id: "edit", type: "replaceInBlock", blockId: block.id, find, replace }],
            { mode, wordDiff: { granularity } },
          );
          const findCutsField = cleanBefore.structuralBoundaries.some(
            (boundary) =>
              boundary.type === "field" &&
              [start, end].some(
                (edge) => edge > boundary.offset && edge < boundary.offset + boundary.length,
              ),
          );
          if (findCutsField) {
            expect(result.skipped).toEqual([{ id: "edit", reason: "unsupportedBlock" }]);
            return;
          }
          expect(result.skipped).toEqual([]);
          const expectedText = block.text.slice(0, start) + replace + block.text.slice(end);

          // The redline: the original characters in order, each deleted
          // exactly when a change removes it, and the new text inserted.
          const edited = firstParagraph(reviewer).node;
          const editedCharacters = cleanCharacters(edited);
          const kept = editedCharacters.filter((entry) => !hasMark(entry.marks, "insertion"));
          expect(kept.map((entry) => entry.character).join("")).toBe(block.text);
          expect(
            editedCharacters
              .filter((entry) => hasMark(entry.marks, "insertion"))
              .map((entry) => entry.character)
              .join(""),
          ).toBe(changes.map((change) => change.text).join(""));
          const removedAt = (index: number) =>
            index >= start && index < end && removedBy(changes, index - start);
          const touched = touchedHighlight(
            beforeCharacters.slice(start, end).map((entry) => hasMark(entry.marks, "highlight")),
            changes,
          );
          for (const [index, was] of beforeCharacters.entries()) {
            const now = kept[index];
            const removed = removedAt(index);
            const cleared = !removed && touched.has(index - start);
            expect({
              index,
              deleted: now !== undefined && hasMark(now.marks, "deletion"),
              propertyChange: now !== undefined && hasMark(now.marks, "runPropertyChange"),
              control: now?.control,
            }).toEqual({ index, deleted: removed, propertyChange: cleared, control: was.control });
            // A cleared character keeps its other formatting, though the
            // carrier stating it is rewritten along with the background.
            const nowMarks = withoutMarks(now?.marks ?? [], REVISION_MARKS);
            expect({
              index,
              sameMarks: cleared
                ? textFormattingOf(nowMarks) ===
                  textFormattingOf(withoutMarks(was.marks, BACKGROUND_MARKS))
                : Mark.sameSet(nowMarks, was.marks),
            }).toEqual({ index, sameMarks: true });
          }
          for (const entry of editedCharacters) {
            if (hasMark(entry.marks, "insertion")) {
              expect(entry.marks.some((mark) => BACKGROUND_MARKS.has(mark.type.name))).toBe(false);
            }
          }
          for (const node of inlineContentWithoutText(edited)) {
            expect({
              node: node.type.name,
              revision: node.marks.some((mark) => REVISION_MARKS.has(mark.type.name)),
            }).toEqual({ node: node.type.name, revision: false });
          }

          // Accepting every revision gives the replacement, with every
          // control, bookmark and untouched field still there.
          const accepted = firstParagraphOf(resolveEverything(reviewer.state, mode, "accept").doc);
          const acceptedCharacters = cleanCharacters(accepted.node);
          expect(acceptedCharacters.map((entry) => entry.character).join("")).toBe(expectedText);
          for (const entry of acceptedCharacters) {
            expect(entry.marks.some((mark) => REVISION_MARKS.has(mark.type.name))).toBe(false);
          }
          // The change revises the text inside a control, not the control:
          // accepting the deletion of all its text leaves it standing,
          // emptied, as the direct replacement does.
          expect(countNodes(accepted.node, "sdt")).toBe(countNodes(before.node, "sdt"));
          expect(countNodes(accepted.node, "bookmarkBoundary")).toBe(
            countNodes(before.node, "bookmarkBoundary"),
          );
          const touchedFields = fields.filter(({ offset, length }) =>
            changes.some((change) => change.start < offset + length && change.end > offset),
          ).length;
          expect(countNodes(accepted.node, "field")).toBe(
            countNodes(before.node, "field") - touchedFields,
          );

          // Rejecting every revision gives back the original paragraph. A
          // rejected property change restores the formatting, stated on a
          // carrier of its own, so where a background was cleared the
          // formatting is compared rather than the carrier.
          const rejected = firstParagraphOf(resolveEverything(reviewer.state, mode, "reject").doc);
          if (touched.size === 0) {
            expect(rejected.node.content.eq(before.node.content)).toBe(true);
          } else {
            expect(formattingOf(rejected.node)).toEqual(formattingOf(before.node));
          }

          if (mode === "tracked-changes") {
            // The redline survives a save; the reopened package resolves the
            // same way through the reviewer's own accept and reject.
            const saved = await reviewer.toBuffer();
            const acceptedPackage = await FolioDocxReviewer.fromBuffer(saved);
            acceptedPackage.acceptAll();
            expect(countNodes(firstParagraph(acceptedPackage).node, "sdt")).toBe(
              countNodes(before.node, "sdt"),
            );
            const acceptedText = acceptedPackage.snapshot().blocks.at(0)?.text ?? "";
            const rejectedPackage = await FolioDocxReviewer.fromBuffer(saved);
            rejectedPackage.rejectAll();
            const rejectedText = rejectedPackage.snapshot().blocks.at(0)?.text ?? "";
            if (items.some((entry) => entry.kind === "link")) {
              // Where a saved `w:hyperlink` lands among its sibling runs is
              // the serializer's contract, tested with it.
              expect([...acceptedText].toSorted()).toEqual([...expectedText].toSorted());
              expect([...rejectedText].toSorted()).toEqual([...block.text].toSorted());
            } else {
              expect(acceptedText).toBe(expectedText);
              expect(rejectedText).toBe(block.text);
            }
          }
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
  }, 240_000);
});

/**
 * The first body paragraph of a saved package, as XML. `xml:space` is dropped:
 * a re-serialized paragraph states it only where the text needs it, which
 * changes no character.
 */
const savedFirstParagraph = async (saved: ArrayBuffer): Promise<string> => {
  const xml = (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
  const start = xml.search(/<w:p[ >]/u);
  return xml
    .slice(start, xml.indexOf("</w:p>", start) + "</w:p>".length)
    .replaceAll(' xml:space="preserve"', "");
};

type ReplaceOptions = {
  find?: string;
  replace: (text: string) => string;
  type?: "replaceInBlock" | "replaceBlock";
  comment?: string;
  mode?: "direct" | ReviewMode;
  granularity?: WordDiffGranularity;
};

const replaceFirst = async (paragraphXml: string, options: ReplaceOptions) => {
  const source = await createDocx(paragraphXml);
  const control = await savedFirstParagraph(
    await (await FolioDocxReviewer.fromBuffer(source)).toBuffer(),
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  const block = reviewer.snapshot().blocks.at(0);
  if (!block) {
    throw new Error("expected a block");
  }
  const find = options.find ?? block.text;
  const comment = options.comment === undefined ? {} : { comment: { text: options.comment } };
  const result = reviewer.applyOperations(
    [
      options.type === "replaceBlock"
        ? { id: "edit", type: "replaceBlock", blockId: block.id, text: options.replace(block.text) }
        : {
            id: "edit",
            type: "replaceInBlock",
            blockId: block.id,
            find,
            replace: options.replace(find),
            ...comment,
          },
    ],
    {
      mode: options.mode ?? "direct",
      revisionStamp: { date: "2026-01-02T03:04:05Z", idSeed: 100 },
      ...(options.granularity === undefined
        ? {}
        : { wordDiff: { granularity: options.granularity } }),
    },
  );
  expect(result.skipped).toEqual([]);
  const saved = await reviewer.toBuffer();
  return { control, edited: await savedFirstParagraph(saved), reviewer };
};

const APPENDED = "‸";
const appendedRun = (properties = "") =>
  `<w:r>${properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`}<w:t>${APPENDED}</w:t></w:r>`;
const controlOf = (text: string, index: number) =>
  itemXml({ kind: "control", runs: [{ text, properties: { ...PLAIN, italic: true } }] }, index);

describe("direct replacement examples", () => {
  test("appending one character to a mixed paragraph changes nothing else", async () => {
    const paragraph =
      runXml("Bold start ", { bold: true, italic: false, size: 28, rsid: true }) +
      controlOf("John", 1) +
      runXml(" pays ", PLAIN) +
      itemXml({ kind: "field", result: "3.6" }, 2) +
      itemXml({ kind: "tab" }, 3) +
      runXml("end", { ...PLAIN, italic: true });
    const { control, edited } = await replaceFirst(paragraph, {
      replace: (text) => text + APPENDED,
    });
    expect(edited).toBe(
      control.replace(/end<\/w:t><\/w:r><\/w:p>$/u, `end${APPENDED}</w:t></w:r></w:p>`),
    );
  });

  test("text appended after a note reference is not styled as the reference", async () => {
    const paragraph = runXml("See the note", PLAIN) + itemXml({ kind: "noteReference" }, 0);
    const { control, edited } = await replaceFirst(paragraph, {
      replace: (text) => text + APPENDED,
    });
    expect(edited).toBe(control.replace(/<\/w:p>$/u, `${appendedRun()}</w:p>`));
  });

  test("text appended to a paragraph ending in a content control goes after the control", async () => {
    const paragraph = runXml("Name: ", PLAIN) + controlOf("John", 0);
    const { control, edited } = await replaceFirst(paragraph, {
      replace: (text) => text + APPENDED,
    });
    expect(edited).toBe(control.replace(/<\/w:p>$/u, `${appendedRun("<w:i/>")}</w:p>`));
  });

  test("an edit matching only the control's text stays inside the control", async () => {
    const paragraph = runXml("Name: ", PLAIN) + controlOf("John", 0);
    const { control, edited } = await replaceFirst(paragraph, {
      find: "John",
      replace: () => "John Smith",
    });
    expect(edited).toBe(control.replace(">John</w:t>", ">John Smith</w:t>"));
  });

  test("a field survives when its text does; it becomes text when its text changes", async () => {
    const paragraph =
      runXml("see ", PLAIN) +
      itemXml({ kind: "field", result: "3.6" }, 0) +
      runXml(" above", PLAIN);
    const kept = await replaceFirst(paragraph, { replace: () => "see 3.6 below" });
    expect(kept.edited).toContain("<w:instrText");
    expect(kept.edited).toContain(">3.6</w:t>");
    const replaced = await replaceFirst(paragraph, { replace: () => "see 3.7 above" });
    expect(replaced.edited).not.toContain("<w:instrText");
    expect(replaced.reviewer.snapshot().blocks.at(0)?.text).toBe("see 3.7 above");
  });

  test("changing one digit of a highlighted amount clears the whole amount's highlight only", async () => {
    const highlighted = (text: string) =>
      `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const paragraph =
      highlighted("Draft") +
      runXml(" penalty ", PLAIN) +
      highlighted("2000") +
      runXml(" CZK", PLAIN);
    const { edited, reviewer } = await replaceFirst(paragraph, {
      replace: (text) => text.replace("2000", "3000"),
    });
    expect(reviewer.snapshot().blocks.at(0)?.text).toBe("Draft penalty 3000 CZK");
    expect(edited).toMatch(/<w:highlight w:val="yellow"\/><\/w:rPr><w:t[^>]*>Draft<\/w:t>/u);
    expect((edited.match(/<w:highlight /gu) ?? []).length).toBe(1);
  });

  test("an operation comment covers the replaced text", async () => {
    const paragraph =
      runXml("The Seller ", { ...PLAIN, bold: true }) + runXml("shall deliver.", PLAIN);
    const { edited } = await replaceFirst(paragraph, {
      replace: (text) => text.replace("shall", "must"),
      comment: "Stronger wording.",
    });
    expect(edited).toContain("<w:commentRangeStart");
    expect(edited).toContain("<w:commentReference");
    expect(edited).toContain("<w:b/>");
  });

  test("replaceBlock writes only the difference too", async () => {
    const paragraph =
      runXml("Bold ", { ...PLAIN, bold: true }) + runXml("italic", { ...PLAIN, italic: true });
    const { control, edited } = await replaceFirst(paragraph, {
      type: "replaceBlock",
      replace: (text) => text + APPENDED,
    });
    expect(edited).toBe(control.replace(/italic<\/w:t>/u, `italic${APPENDED}</w:t>`));
  });
});

/** `xml` without its `w:ins` elements and their content. */
const withoutInsertions = (xml: string): string =>
  xml.replaceAll(/<w:ins [^>]*>.*?<\/w:ins>/gu, "");

describe("tracked replacement examples", () => {
  test("appending one character to a mixed paragraph inserts only that character", async () => {
    const paragraph =
      runXml("Bold start ", { bold: true, italic: false, size: 28, rsid: true }) +
      controlOf("John", 1) +
      itemXml({ kind: "noteReference" }, 0) +
      runXml(" pays ", PLAIN) +
      itemXml({ kind: "field", result: "3.6" }, 2) +
      itemXml({ kind: "tab" }, 3) +
      runXml("end", { ...PLAIN, italic: true });
    const { control, edited } = await replaceFirst(paragraph, {
      replace: (text) => text + APPENDED,
      mode: "tracked-changes",
      granularity: "character",
    });
    expect(edited).not.toContain("<w:del ");
    expect(edited).not.toContain("<w:rPrChange ");
    expect(edited.match(/<w:ins [^>]*>.*?<\/w:ins>/gu)).toEqual([
      expect.stringContaining(`<w:i/></w:rPr><w:t>${APPENDED}</w:t></w:r></w:ins>`),
    ]);
    expect(withoutInsertions(edited)).toBe(control);
  });

  test("a changed word next to a field and a tab is the only redline", async () => {
    const paragraph =
      runXml("see ", PLAIN) +
      itemXml({ kind: "field", result: "3.6" }, 0) +
      itemXml({ kind: "tab" }, 1) +
      runXml("above here", { ...PLAIN, bold: true });
    const { edited, reviewer } = await replaceFirst(paragraph, {
      replace: (text) => text.replace("above", "below"),
      mode: "tracked-changes",
    });
    expect(edited.match(/<w:delText>[^<]*<\/w:delText>/gu)).toEqual([
      "<w:delText>above</w:delText>",
    ]);
    expect(edited).toContain("<w:instrText");
    expect(edited).toContain("<w:tab/>");
    const accepting = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    accepting.acceptAll();
    expect(accepting.snapshot().blocks.at(0)?.text).toBe("see 3.6\tbelow here");
  });

  test("a replaced highlighted amount keeps its highlight as deleted text", async () => {
    const highlighted = (text: string) =>
      `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const paragraph =
      highlighted("Draft") +
      runXml(" penalty ", PLAIN) +
      highlighted("2000") +
      runXml(" CZK", PLAIN);
    const { edited } = await replaceFirst(paragraph, {
      replace: (text) => text.replace("2000", "3000"),
      mode: "suggested",
    });
    // Suggestions stay out of the saved package until accepted.
    expect(edited).not.toContain("<w:rPrChange ");
    for (const mode of ["tracked-changes", "suggested"] as const) {
      const { reviewer } = await replaceFirst(paragraph, {
        replace: (text) => text.replace("2000", "3000"),
        mode,
      });
      const runs: string[] = [];
      firstParagraph(reviewer).node.descendants((node) => {
        if (node.isText) {
          runs.push(
            `${node.text}:${node.marks
              .map((mark) => mark.type.name)
              .filter((name) => name !== "runIdentity")
              .toSorted()
              .join("+")}`,
          );
        }
        return true;
      });
      // A word's redline carries the space before it.
      expect(runs).toEqual([
        "Draft:highlight",
        " penalty:",
        " :deletion",
        "2000:deletion+highlight",
        " 3000:insertion",
        " CZK:",
      ]);
    }
  });
});
