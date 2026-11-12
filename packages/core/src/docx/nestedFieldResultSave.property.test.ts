/**
 * Complex fields nested in other fields, at any depth, in the code or in the
 * result, survive a save that writes the paragraph from the model: the
 * reader sees the same text, and the saved paragraph holds the same field
 * characters and instructions in the same order.
 *
 * A save writes a paragraph from the model once anything structural changed
 * in the document, such as a paragraph inserted elsewhere. A citation or a
 * cross-reference field nested in another field's result lost its own
 * `separate` and `end` there, and with them its result text.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { RELATIONSHIP_TYPES } from "./relsParser";

setDefaultTimeout(propertyTestTimeout(160_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

type Node =
  | { kind: "text"; text: string }
  | { kind: "field"; instruction: string; code: Node[]; result: Node[] };

const WORDS = ["alpha", "beta", "gamma", "delta"] as const;

const nodeArbitrary: fc.Memo<Node> = fc.memo((depth) =>
  depth <= 1
    ? fc.record({
        kind: fc.constant("text" as const),
        text: fc.constantFrom(...WORDS),
      })
    : fc.oneof(
        { depthSize: "small" },
        fc.record({
          kind: fc.constant("text" as const),
          text: fc.constantFrom(...WORDS),
        }),
        fc.record({
          kind: fc.constant("field" as const),
          instruction: fc.constantFrom(
            " ADDIN EN.CITE ",
            " REF _Ref1 \\h ",
            " QUOTE ",
            " DOCPROPERTY Title ",
          ),
          code: fc.array(nodeArbitrary(depth - 1), { maxLength: 2 }),
          result: fc.array(nodeArbitrary(depth - 1), {
            minLength: 1,
            maxLength: 3,
          }),
        }),
      ),
);

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const fieldChar = (type: string) => run(`<w:fldChar w:fldCharType="${type}"/>`);

/** The node's markup; text in a code region is instruction text. */
const xmlOf = (node: Node, inCode: boolean): string => {
  if (node.kind === "text")
    return inCode
      ? run(`<w:instrText xml:space="preserve"> ${node.text} </w:instrText>`)
      : run(`<w:t xml:space="preserve">${node.text} </w:t>`);
  return (
    fieldChar("begin") +
    run(`<w:instrText xml:space="preserve">${node.instruction}</w:instrText>`) +
    node.code.map((child) => xmlOf(child, true)).join("") +
    fieldChar("separate") +
    node.result.map((child) => xmlOf(child, inCode)).join("") +
    fieldChar("end")
  );
};

/** What a reader shows: a field shows its result, never its code. */
const visibleOf = (node: Node, inCode: boolean): string => {
  if (node.kind === "text") return inCode ? "" : `${node.text} `;
  return node.result.map((child) => visibleOf(child, inCode)).join("");
};

const documentOf = async (nodes: readonly Node[]): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="11111111">${run('<w:t xml:space="preserve">First.</w:t>')}</w:p>` +
      `<w:p w14:paraId="22222222">${run('<w:t xml:space="preserve">Lead </w:t>')}` +
      nodes.map((node) => xmlOf(node, false)).join("") +
      `${run('<w:t xml:space="preserve">tail.</w:t>')}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** The field characters and instructions of the cited paragraph, in order. */
const fieldStructureOf = async (bytes: ArrayBuffer): Promise<string[]> => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const start = xml.indexOf('w14:paraId="22222222"');
  const paragraph = xml.slice(start, xml.indexOf("</w:p>", start));
  return [
    ...paragraph.matchAll(
      /<w:fldChar w:fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)<\/w:instrText>/gu,
    ),
  ].map(([, type, instruction]) => type ?? `instr:${instruction ?? ""}`);
};

const paragraphText = (reviewer: FolioDocxReviewer): string | undefined =>
  reviewer.getContent().find(({ id }) => id === "22222222")?.text;

describe("nested complex fields across a save that writes from the model", () => {
  test("keep their text and their structure at any depth", async () => {
    const law = (counts: { nested: number }) =>
      fc.asyncProperty(
        fc.array(nodeArbitrary(5), { minLength: 1, maxLength: 3 }),
        async (nodes) => {
          const source = await documentOf(nodes);
          const expected = `Lead ${nodes.map((node) => visibleOf(node, false)).join("")}tail.`;
          const reviewer = await FolioDocxReviewer.fromBuffer(source);
          expect(paragraphText(reviewer)).toBe(expected);
          // A paragraph inserted elsewhere makes the save write the body from
          // the model.
          const result = reviewer.applyOperations(
            [
              {
                id: "insert",
                type: "insertAfterBlock",
                blockId: "11111111",
                text: "Added.",
              },
            ],
            { undefinedReferences: "refuse", mode: "direct" },
          );
          expect(result.applied).toHaveLength(1);
          const saved = await reviewer.toBuffer();
          expect(await fieldStructureOf(saved)).toEqual(
            await fieldStructureOf(source),
          );
          const reopened = await FolioDocxReviewer.fromBuffer(saved);
          expect(paragraphText(reopened)).toBe(expected);
          // An edit of text a field shows either applies or is refused, and
          // leaves every field's characters as they were.
          for (const mode of ["direct", "tracked-changes"] as const) {
            const editing = await FolioDocxReviewer.fromBuffer(saved);
            const edit = editing.applyOperations(
              [
                {
                  id: "edit",
                  type: "replaceInBlock",
                  blockId: "22222222",
                  find: WORDS.find((word) => expected.includes(word)) ?? "tail",
                  replace: "omega",
                },
              ],
              { undefinedReferences: "refuse", mode },
            );
            const edited = await editing.toBuffer();
            const editedReader = await FolioDocxReviewer.fromBuffer(edited);
            if (mode === "tracked-changes") editedReader.acceptAll();
            if (edit.applied.length === 0)
              expect(paragraphText(editedReader)).toBe(expected);
            else expect(paragraphText(editedReader)).toContain("omega");
            expect(await fieldStructureOf(edited)).toEqual(
              await fieldStructureOf(source),
            );
          }
          if (
            nodes.some(
              (node) =>
                node.kind === "field" &&
                [...node.code, ...node.result].some(
                  ({ kind }) => kind === "field",
                ),
            )
          )
            counts.nested += 1;
        },
      );
    await assertProperty(law({ nested: 0 }), {
      id: "nested complex fields survive a model save",
      numRuns: 150,
    });
    // A fixed sample reaches fields nested in fields. Each case saves and
    // reopens the document several times, so the sample is a third of the runs.
    const RUNS = 50;
    const counts = { nested: 0 };
    await checkCoverageSample(law(counts), RUNS);
    expect(counts.nested).toBeGreaterThan(0);
  });
});
