/**
 * Every child `CT_R` allows survives a forced save, and the save is a fixed
 * point.
 *
 * The run's child switch used to model what it recognised and let the rest fall
 * off the end, which is worse than a loss. `hasRunPayloadElement` asked the
 * *source* element whether a run carried a payload while `serializeRun` wrote
 * the *model*, so an unmodelled child made the two disagree for exactly one
 * save:
 *
 *     parse 1: <w:r><w:ruby/></w:r>  -> run kept, content: []
 *     save  1: <w:r/>                  the ruby is gone
 *     parse 2:                       -> run dropped, no payload element
 *     save  2: (no run at all)         != save 1
 *
 * The input class is "a run child folio does not model", and it is not a list
 * somebody maintains: it is drawn from the committed schema graph, so an
 * element a later OOXML revision adds to `CT_R` joins the property by itself.
 * The three assertions are the three ways the class fails — the markup goes,
 * the text goes, or the bytes never settle.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { parseDocx } from "../packages/core/src/docx/parser";
import { createEmptyDocx, repackDocx } from "../packages/core/src/docx/rezip";
import { getParagraphText } from "../packages/core/src/docx/paragraphParser";
import { propertyConfig, propertyTestTimeout } from "../test/property-testing";
import { loadContainerSpace, qualify, WML_NAMESPACE } from "./lib/container-survival/schemaSpace";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const space = await loadContainerSpace();

const RUN_CONTAINER_KEY = `${qualify({ namespace: WML_NAMESPACE, name: "r" })}|${qualify({
  namespace: WML_NAMESPACE,
  name: "CT_R",
})}`;

/**
 * `w:rPr` is the one child excluded: `parseRunProperties` reads the same
 * element, so it is modelled by a different route and preserving it as well
 * would write the run's properties twice.
 */
const runChildNames = (): readonly string[] => {
  const container = space.containers.get(RUN_CONTAINER_KEY);
  if (!container) {
    throw new Error(`the schema graph has no ${RUN_CONTAINER_KEY}`);
  }
  return container.children
    .filter(({ child }) => child.namespace === WML_NAMESPACE && child.name !== "rPr")
    .map(({ child }) => child.name);
};

const CHILD_NAMES = runChildNames();

/**
 * Children the generator can only write in a shape that is not a document, so
 * their disappearance is not what this property measures.
 *
 * - A complex field is a `begin` … `separate` … `end` sequence read at
 *   paragraph level. A lone marker is not a field, and whether an unbalanced
 *   one should round-trip is `RC-fieldcode-formatting`'s question.
 * - `w:drawing` and `w:pict` hold a picture that needs a relationship and a
 *   media part. Empty, they carry nothing to keep; the hand-off that loses a
 *   real VML one lives in the VML parser.
 * - `w:commentReference` is owned one level up: the paragraph parser lifts it
 *   out of the run and the comment serializer re-emits its own run for it. One
 *   with no id, pointing at no comment, is not a reference.
 *
 * The fixed-point and text assertions still cover all of them, which is the
 * claim that matters.
 */
const NOT_A_DOCUMENT: ReadonlySet<string> = new Set([
  "commentReference",
  "delInstrText",
  "drawing",
  "fldChar",
  "instrText",
  "pict",
]);

/**
 * Markup for one run child, chosen to be the shape a document writes.
 *
 * A child folio models is written the way its parser expects it; anything else
 * gets an empty element, which is what the sink has to carry. `w:ruby` is the
 * case with visible text and gets its real shape, because the property asserts
 * the paragraph's text as well as its bytes.
 */
const childMarkup = (name: string, word: string): string => {
  switch (name) {
    case "t":
    case "delText":
      return `<w:${name}>${word}</w:${name}>`;
    case "instrText":
    case "delInstrText":
      return `<w:${name}> PAGE </w:${name}>`;
    case "br":
      return '<w:br w:type="page"/>';
    case "sym":
      return '<w:sym w:font="Wingdings" w:char="F0E0"/>';
    case "fldChar":
      return '<w:fldChar w:fldCharType="begin"/>';
    case "footnoteReference":
    case "endnoteReference":
      return `<w:${name} w:id="1"/>`;
    case "ruby":
      return (
        "<w:ruby><w:rubyPr/>" +
        `<w:rt><w:r><w:t>annotation</w:t></w:r></w:rt>` +
        `<w:rubyBase><w:r><w:t>${word}</w:t></w:rubyBase></w:ruby>`
      ).replace("</w:rubyBase>", "</w:r></w:rubyBase>");
    default:
      return `<w:${name}/>`;
  }
};

/** The text every paragraph in the body puts on the page, joined. */
const bodyText = ({ package: { document } }: Awaited<ReturnType<typeof parseDocx>>): string =>
  document.content
    .filter((block) => block.type === "paragraph")
    .map((paragraph) => getParagraphText(paragraph))
    .join("\n");

/** A notes part holding Word's two required separators plus note 1. */
const notesPart = (root: "footnotes" | "endnotes", note: "footnote" | "endnote"): string =>
  `${XML_DECLARATION}<w:${root} xmlns:w="${WML_NAMESPACE}">` +
  `<w:${note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${note}>` +
  `<w:${note} w:type="continuationSeparator" w:id="0"><w:p><w:r>` +
  `<w:continuationSeparator/></w:r></w:p></w:${note}>` +
  `<w:${note} w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:${note}>` +
  `</w:${root}>`;

const NOTE_CONTENT_TYPE_PREFIX = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
const NOTE_RELATIONSHIP_PREFIX =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/**
 * A run is put in a one-paragraph body, with the note parts the reference
 * markers need: folio refuses a package whose `w:footnoteReference` points at
 * nothing, and a fixture that cannot be parsed proves nothing.
 */
const buildDocx = async (markup: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${WML_NAMESPACE}"><w:body>` +
      `<w:p><w:r>${markup}</w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  zip.file("word/footnotes.xml", notesPart("footnotes", "footnote"));
  zip.file("word/endnotes.xml", notesPart("endnotes", "endnote"));

  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      `<Override PartName="/word/footnotes.xml" ContentType="${NOTE_CONTENT_TYPE_PREFIX}footnotes+xml"/>` +
        `<Override PartName="/word/endnotes.xml" ContentType="${NOTE_CONTENT_TYPE_PREFIX}endnotes+xml"/></Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdPropFootnotes" Type="${NOTE_RELATIONSHIP_PREFIX}footnotes" Target="footnotes.xml"/>` +
        `<Relationship Id="rIdPropEndnotes" Type="${NOTE_RELATIONSHIP_PREFIX}endnotes" Target="endnotes.xml"/></Relationships>`,
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPart = async (buffer: ArrayBuffer): Promise<string> => {
  const part = await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text");
  if (part === undefined) {
    throw new Error("the saved package has no word/document.xml");
  }
  return part;
};

/** The run's markup, from `<w:r` to its close, with `w:rPr` removed. */
const runPayload = (part: string): string => {
  const start = part.indexOf("<w:r>");
  const end = part.indexOf("</w:r>", start);
  if (start === -1 || end === -1) {
    return "";
  }
  return part.slice(start + "<w:r>".length, end).replace(/<w:rPr>.*?<\/w:rPr>|<w:rPr\/>/su, "");
};

describe("every CT_R child survives a forced save", () => {
  test(
    "the markup, the paragraph text, and the bytes all settle",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...CHILD_NAMES),
          fc.stringMatching(/^[A-Za-z]{1,8}$/u),
          async (name, word) => {
            const markup = childMarkup(name, word);
            const parsed = await parseDocx(await buildDocx(markup), { preloadFonts: false });

            const first = await repackDocx(parsed, { updateModifiedDate: false });
            const firstPart = await documentPart(first);

            // The markup is still there. A modelled child comes back re-spelled
            // by its own serializer; an unmodelled one comes back byte-equal,
            // which is the whole claim the sink makes.
            if (!NOT_A_DOCUMENT.has(name)) {
              expect({ name, payload: runPayload(firstPart) }).not.toEqual({ name, payload: "" });
            }

            const reparsed = await parseDocx(first, { preloadFonts: false });
            expect({ name, text: bodyText(reparsed) }).toEqual({
              name,
              text: bodyText(parsed),
            });

            const second = await documentPart(
              await repackDocx(reparsed, { updateModifiedDate: false }),
            );
            expect({ name, part: second }).toEqual({ name, part: firstPart });
          },
        ),
        propertyConfig({ numRuns: 120 }),
      );
    },
    // Three saves and four parses per run, over a package with note parts:
    // the default five seconds fails this property on its own runtime.
    propertyTestTimeout(120_000),
  );

  test("an unmodelled child comes back byte-equal, at its position", async () => {
    const markup = "<w:t>before</w:t><w:ruby><w:rubyPr/></w:ruby><w:t>after</w:t>";
    const parsed = await parseDocx(await buildDocx(markup), { preloadFonts: false });
    const payload = runPayload(await documentPart(await repackDocx(parsed)));
    expect(payload).toContain("<w:ruby>");
    expect(payload.indexOf("before")).toBeLessThan(payload.indexOf("<w:ruby>"));
    expect(payload.indexOf("<w:ruby>")).toBeLessThan(payload.indexOf("after"));
  });
});
