/**
 * Accept All and Reject All against a revision-resolution reference.
 *
 * Each fixture is a small package with pending revisions: paragraph marks
 * deleted or inserted between paragraphs whose properties differ, with and
 * without pending property changes or tracked content, chains of them,
 * nested revisions by two authors, and property changes; some are shapes a
 * reviewer wrote with tracking on. Its references are the paragraphs a
 * reference implementation left after accepting, and after rejecting, every
 * revision: text, paragraph boundaries, identity (`w14:paraId`), paragraph
 * properties, the paragraph mark's run properties and run formatting.
 *
 * The rule they pin: a paragraph mark that resolution removes (a deletion
 * accepted, an insertion rejected) takes its paragraph's properties with it.
 * The paragraph left is the one whose mark stays, the next one, whatever
 * the first still holds: its identity, style, numbering, alignment, spacing,
 * indentation and mark run properties survive, and the first's runs keep
 * only their own direct formatting. The result does not depend on the order
 * the revisions are resolved in.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { getLocalName, parseXml, type XmlElement } from "../docx/xmlParser";
import {
  createHarnessState,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
} from "./editorHarness";

type ParagraphSummary = {
  text: string;
  paraId: string | null;
  pPr: string;
  markRPr: string;
  runs?: string[];
};
type BlockSummary = ParagraphSummary | { table: BlockSummary[][][] };

type NoteSummaries = Partial<Record<"footnote" | "endnote", Record<string, BlockSummary[]>>>;

type ReferenceFixture = {
  origin: "constructed" | "authored";
  body: string;
  footnotes?: string;
  endnotes?: string;
  accept?: BlockSummary[];
  reject?: BlockSummary[];
  /** The notes left: a note goes with its reference. */
  acceptNotes?: NoteSummaries;
  rejectNotes?: NoteSummaries;
  /** Every revision this author made rejected, and nothing else resolved. */
  rejectAuthor?: string;
  rejectAuthorResult?: BlockSummary[];
};

type Reference = {
  package: { styles: string; numbering: string; settings: string; sectPr: string };
  fixtures: Record<string, ReferenceFixture>;
};

const reference = JSON.parse(
  readFileSync(
    path.join(import.meta.dir, "__fixtures__", "revision-resolution-reference.json"),
    "utf8",
  ),
) as Reference;

// ============================================================================
// PACKAGE
// ============================================================================

const NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'mc:Ignorable="w14"';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const NOTE_SEPARATORS = (kind: "footnote" | "endnote") =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

const buildPackage = async (fixture: ReferenceFixture): Promise<Uint8Array> => {
  const zip = new JSZip();
  const notes = (["footnote", "endnote"] as const).filter((kind) => fixture[`${kind}s`]);
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${MAIN}.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="${MAIN}.numbering+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${MAIN}.settings+xml"/>` +
      notes
        .map(
          (kind) => `<Override PartName="/word/${kind}s.xml" ContentType="${MAIN}.${kind}s+xml"/>`,
        )
        .join("") +
      `</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIPS}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="${RELATIONSHIPS}/numbering" Target="numbering.xml"/>` +
      `<Relationship Id="rId3" Type="${RELATIONSHIPS}/settings" Target="settings.xml"/>` +
      notes
        .map(
          (kind, index) =>
            `<Relationship Id="rId${10 + index}" Type="${RELATIONSHIPS}/${kind}s" Target="${kind}s.xml"/>`,
        )
        .join("") +
      `</Relationships>`,
  );
  const body = fixture.body.includes("<w:sectPr")
    ? fixture.body
    : `${fixture.body}${reference.package.sectPr}`;
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${body}</w:body></w:document>`,
  );
  zip.file("word/styles.xml", reference.package.styles);
  zip.file("word/numbering.xml", reference.package.numbering);
  zip.file("word/settings.xml", reference.package.settings);
  for (const kind of notes) {
    zip.file(
      `word/${kind}s.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s ${NAMESPACES}>${NOTE_SEPARATORS(kind)}${fixture[`${kind}s`] ?? ""}</w:${kind}s>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

// ============================================================================
// SUMMARY (the reference's canonical form)
// ============================================================================

/** Children that carry no resolution result: fonts, languages and marks a save adds or drops. */
const IGNORED_ELEMENTS = new Set([
  "rFonts",
  "lang",
  "szCs",
  "bCs",
  "iCs",
  "kern",
  "noProof",
  "rsid",
  "proofErr",
  "lastRenderedPageBreak",
  "bookmarkStart",
  "bookmarkEnd",
]);
const MARK_REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo"]);

const byCodePoint = (left: string, right: string): number => {
  if (left < right) return -1;
  return left > right ? 1 : 0;
};

/** Deleted text reads `[-…-]` and inserted text `[+…+]`; resolved documents hold neither. */
const withRevisionBrackets = (text: string, mode: "ins" | "del" | null): string => {
  if (mode === "del") return `[-${text}-]`;
  return mode === "ins" ? `[+${text}+]` : text;
};

const children = (element: XmlElement): XmlElement[] =>
  (element.elements ?? []).filter((child) => child.type === "element");

const isIgnoredAttribute = (name: string): boolean => {
  const local = getLocalName(name);
  return (
    name.startsWith("w14:") ||
    name.toLowerCase().includes("rsid") ||
    local === "id" ||
    local === "date"
  );
};

/** Compact canonical XML of an element's children: local names, sorted attributes. */
const canonical = (elements: readonly XmlElement[]): string =>
  elements
    .filter((child) => !IGNORED_ELEMENTS.has(getLocalName(child.name)))
    .map((child) => {
      const name = getLocalName(child.name);
      const attributes = Object.entries(child.attributes ?? {})
        .filter(([key]) => !isIgnoredAttribute(key))
        .map(([key, value]) => [getLocalName(key), String(value)] as const)
        .toSorted(([left], [right]) => byCodePoint(left, right))
        .map(([key, value]) => ` ${key}="${value}"`)
        .join("");
      const inner = canonical(children(child));
      return inner ? `<${name}${attributes}>${inner}</${name}>` : `<${name}${attributes}/>`;
    })
    .join("");

const textOf = (element: XmlElement): string =>
  (element.elements ?? []).map((child) => String(child.text ?? "")).join("");

const runText = (run: XmlElement): string =>
  children(run)
    .map((child) => {
      const name = getLocalName(child.name);
      // Field instructions are not text a reader sees: a link restored as a
      // field or as a hyperlink element reads the same.
      if (name === "t" || name === "delText") return textOf(child);
      if (name === "tab") return "\t";
      if (name === "br" || name === "cr") return "\n";
      if (name === "footnoteReference" || name === "endnoteReference") {
        return `<${name}:${String(child.attributes?.["w:id"] ?? "")}>`;
      }
      return "";
    })
    .join("");

const summarizeParagraph = (paragraph: XmlElement): ParagraphSummary => {
  const properties = children(paragraph).find((child) => getLocalName(child.name) === "pPr");
  const own = (properties ? children(properties) : []).filter(
    (child) => !["rPr", "pPrChange", "sectPr"].includes(getLocalName(child.name)),
  );
  const markRun = (properties ? children(properties) : []).find(
    (child) => getLocalName(child.name) === "rPr",
  );
  let text = "";
  const runs: [string, string][] = [];
  const walk = (node: XmlElement, mode: "ins" | "del" | null): void => {
    for (const child of children(node)) {
      const name = getLocalName(child.name);
      if (name === "r") {
        const runProperties = children(child).find((inner) => getLocalName(inner.name) === "rPr");
        const formatting = runProperties ? canonical(children(runProperties)) : "";
        const raw = runText(child);
        if (!raw) continue;
        const value = withRevisionBrackets(raw, mode);
        text += value;
        const last = runs.at(-1);
        if (last && last[1] === formatting) last[0] += value;
        else runs.push([value, formatting]);
      } else if (MARK_REVISIONS.has(name)) {
        walk(child, name === "del" || name === "moveFrom" ? "del" : "ins");
      } else if (
        ["hyperlink", "smartTag", "customXml", "sdt", "sdtContent", "fldSimple"].includes(name)
      ) {
        walk(child, mode);
      }
    }
  };
  walk(paragraph, null);
  const summary: ParagraphSummary = {
    text,
    paraId: String(paragraph.attributes?.["w14:paraId"] ?? "") || null,
    pPr: canonical(own),
    markRPr: markRun
      ? canonical(
          children(markRun).filter((child) => !MARK_REVISIONS.has(getLocalName(child.name))),
        )
      : "",
  };
  const formatted = runs.filter(([, formatting]) => formatting).map(([t, f]) => `${t}|${f}`);
  if (formatted.length > 0) summary.runs = formatted;
  return summary;
};

const summarizeBlocks = (parent: XmlElement): BlockSummary[] =>
  children(parent).flatMap((child): BlockSummary[] => {
    const name = getLocalName(child.name);
    if (name === "p") return [summarizeParagraph(child)];
    if (name === "tbl") {
      return [
        {
          table: children(child)
            .filter((row) => getLocalName(row.name) === "tr")
            .map((row) =>
              children(row)
                .filter((cell) => getLocalName(cell.name) === "tc")
                .map(summarizeBlocks),
            ),
        },
      ];
    }
    if (name === "sdt") {
      const content = children(child).find((inner) => getLocalName(inner.name) === "sdtContent");
      return content ? summarizeBlocks(content) : [];
    }
    return [];
  });

const NOTE_KINDS = ["footnote", "endnote"] as const;
const SEPARATOR_TYPES = new Set(["separator", "continuationSeparator", "continuationNotice"]);

const summarizeNotes = async (bytes: Uint8Array | ArrayBuffer): Promise<NoteSummaries> => {
  const zip = await JSZip.loadAsync(bytes);
  const summaries: NoteSummaries = {};
  for (const kind of NOTE_KINDS) {
    const xml = await zip.file(`word/${kind}s.xml`)?.async("text");
    if (!xml) continue;
    const root = children(parseXml(xml))[0];
    const notes: Record<string, BlockSummary[]> = {};
    for (const note of root ? children(root) : []) {
      if (getLocalName(note.name) !== kind) continue;
      if (SEPARATOR_TYPES.has(String(note.attributes?.["w:type"] ?? ""))) continue;
      notes[String(note.attributes?.["w:id"] ?? "")] = summarizeBlocks(note);
    }
    if (Object.keys(notes).length > 0) summaries[kind] = notes;
  }
  return summaries;
};

const summarizeBody = async (bytes: Uint8Array | ArrayBuffer): Promise<BlockSummary[]> => {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("word/document.xml")?.async("text");
  if (!xml) throw new Error("The saved package has no main document part");
  const root = children(parseXml(xml))[0];
  const body = root && children(root).find((child) => getLocalName(child.name) === "body");
  if (!body) throw new Error("The saved main document has no body");
  return summarizeBlocks(body);
};

/** Identity is compared where the reference kept the input's ids; it regenerates a few. */
const notesWithoutRegeneratedIds = (
  notes: NoteSummaries | undefined,
  input: ReadonlySet<string>,
): NoteSummaries =>
  Object.fromEntries(
    Object.entries(notes ?? {}).map(([kind, byId]) => [
      kind,
      Object.fromEntries(
        Object.entries(byId).map(([id, blocks]) => [id, withoutRegeneratedIds(blocks, input)]),
      ),
    ]),
  );

const withoutRegeneratedIds = (
  summary: readonly BlockSummary[],
  input: ReadonlySet<string>,
): BlockSummary[] =>
  summary.map((block) =>
    "table" in block
      ? { table: block.table.map((row) => row.map((cell) => withoutRegeneratedIds(cell, input))) }
      : {
          ...block,
          paraId: block.paraId !== null && input.has(block.paraId) ? block.paraId : null,
        },
  );

const inputParaIds = (body: string): Set<string> =>
  new Set([...body.matchAll(/w14:paraId="([0-9A-Fa-f]{8})"/gu)].map((match) => match[1] ?? ""));

// ============================================================================
// CASES
// ============================================================================

const inputIds = (fixture: ReferenceFixture): Set<string> =>
  inputParaIds(`${fixture.body}${fixture.footnotes ?? ""}${fixture.endnotes ?? ""}`);

describe("tracked-change resolution matches the reference", () => {
  for (const [name, fixture] of Object.entries(reference.fixtures)) {
    for (const mode of ["accept", "reject"] as const) {
      const expected = fixture[mode];
      if (!expected) continue;
      test(`${name} › ${mode} all`, async () => {
        const document = await parseShapeDocument(await buildPackage(fixture));
        const resolved = resolveAllChanges(createHarnessState(document, "editing"), mode);
        const saved = await saveHarnessState(resolved, document);
        const ids = inputIds(fixture);
        expect(withoutRegeneratedIds(await summarizeBody(saved.bytes), ids)).toEqual(
          withoutRegeneratedIds(expected, ids),
        );
      });
      const expectedNotes = fixture[`${mode}Notes`];
      if (!expectedNotes) continue;
      test(`${name} › ${mode} all, every story`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          (await buildPackage(fixture)).slice().buffer,
        );
        if (mode === "accept") reviewer.acceptAll();
        else reviewer.rejectAll();
        const saved = await reviewer.toBuffer();
        const ids = inputIds(fixture);
        expect(withoutRegeneratedIds(await summarizeBody(saved), ids)).toEqual(
          withoutRegeneratedIds(expected, ids),
        );
        expect(notesWithoutRegeneratedIds(await summarizeNotes(saved), ids)).toEqual(
          notesWithoutRegeneratedIds(expectedNotes, ids),
        );
      });
    }
    const { rejectAuthor, rejectAuthorResult } = fixture;
    if (rejectAuthor === undefined || rejectAuthorResult === undefined) continue;
    test(`${name} › reject ${rejectAuthor}'s revisions`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        (await buildPackage(fixture)).slice().buffer,
      );
      for (const change of reviewer.getChanges().filter(({ author }) => author === rejectAuthor)) {
        reviewer.rejectChange(change);
      }
      const ids = inputIds(fixture);
      expect(withoutRegeneratedIds(await summarizeBody(await reviewer.toBuffer()), ids)).toEqual(
        withoutRegeneratedIds(rejectAuthorResult, ids),
      );
    });
  }
});
