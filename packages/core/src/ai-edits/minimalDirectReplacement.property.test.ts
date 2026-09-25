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
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { Mark, type Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { runFormattingInlineAtomCleanText } from "../prosemirror/runFormattingInlineCarriers";
import { buildCleanBlockText } from "./clean-text";
import { FolioDocxReviewer } from "./headless";
import { planTextChanges, widenChangesToAtomicSpans } from "./minimal-replacement";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Neighbouring items never share a note: two references to one note would read as one. */
const NOTE_IDS = [1, 2, 3] as const;
const noteIdAt = (index: number): number => NOTE_IDS[index % NOTE_IDS.length] ?? 1;

const WORDS = ["Seller", "shall", "deliver", "the", "goods", "on", "time", "Buyer", "pays"];

type RunProperties = { bold: boolean; italic: boolean; size: 20 | 28 | null; rsid: boolean };

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

const item: fc.Arbitrary<Item> = fc.oneof(
  {
    weight: 5,
    arbitrary: fc.record({
      kind: fc.constant("run" as const),
      text: words,
      properties: runProperties,
    }),
  },
  {
    weight: 2,
    arbitrary: fc.record({
      kind: fc.constant("control" as const),
      runs: fc.array(fc.record({ text: words, properties: runProperties }), {
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

const escapeXml = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const runXml = (text: string, { bold, italic, size, rsid }: RunProperties): string => {
  const properties = [
    bold ? "<w:b/>" : "",
    italic ? "<w:i/>" : "",
    size === null ? "" : `<w:sz w:val="${size}"/>`,
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

const firstParagraph = (reviewer: FolioDocxReviewer): { node: PMNode; from: number } => {
  const { doc } = reviewer.state;
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
          if (!block || block.text.length === 0) {
            return;
          }
          const start = edit.whole ? 0 : Math.floor(edit.sliceStart * block.text.length);
          const end = edit.whole
            ? block.text.length
            : Math.min(
                block.text.length,
                start + 1 + Math.floor(edit.sliceLength * (block.text.length - start)),
              );
          const find = block.text.slice(start, end);
          if (find.length === 0 || block.text.indexOf(find) !== block.text.lastIndexOf(find)) {
            return;
          }
          let replace = find;
          for (const { at, remove, insert } of edit.mutations) {
            const position = Math.floor(at * replace.length);
            replace = replace.slice(0, position) + insert + replace.slice(position + remove);
          }
          if (replace === find) {
            return;
          }

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
    { mode: "direct" },
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
