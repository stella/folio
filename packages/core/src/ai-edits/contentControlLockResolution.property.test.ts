/**
 * Resolving around a change a content control's lock holds gives the rest of
 * the document what resolving everything at once gives it.
 *
 * With a held change in a story, `acceptAll` and `rejectAll` resolve every
 * other change on its own instead of in the one pass the whole-story resolver
 * makes. For random tracked edits after a paragraph holding a locked control
 * with a pending revision (replacements, deletions, splits, joins across
 * paragraph marks, insertions, alignment changes), the paragraphs after it
 * must come out as they do when the same document, without the lock, is
 * resolved in one pass.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { ensureParaIds } from "../docx/ensureParaIds";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditOperation } from "./types";

setDefaultTimeout(propertyTestTimeout(240_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const LOCK = '<w:lock w:val="contentLocked"/>';

const createDocx = async (bodyXml: string): Promise<Uint8Array> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>${bodyXml}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await ensureParaIds(await zip.generate({ compression: "store" })))
    .docx;
};

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** The paragraph holding the locked control, with its pending insertion. */
const LOCKED_PARAGRAPH =
  `<w:p>${run("Held ")}<w:sdt><w:sdtPr><w:alias w:val="Guarded"/>${LOCK}</w:sdtPr>` +
  `<w:sdtContent>${run("clause")}<w:ins w:id="900" w:author="Author" w:date="2024-01-01T00:00:00Z">${run(" text")}</w:ins></w:sdtContent></w:sdt></w:p>`;

const WORDS = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"] as const;
const ALIGNMENTS = ["left", "center", "right", "both"] as const;

const paragraphXml = (words: readonly string[], alignment: string) =>
  `<w:p><w:pPr><w:jc w:val="${alignment}"/></w:pPr>${run(words.join(" "))}</w:p>`;

type Intent =
  | { kind: "replace"; block: number; word: number; replacement: string }
  | { kind: "delete"; block: number }
  | { kind: "split"; block: number; at: number }
  | { kind: "merge"; block: number }
  | { kind: "insert"; block: number; text: string }
  | { kind: "align"; block: number; alignment: (typeof ALIGNMENTS)[number] };

const block = fc.nat({ max: 7 });
const intentArbitrary: fc.Arbitrary<Intent> = fc.oneof(
  fc.record({
    kind: fc.constant("replace" as const),
    block,
    word: fc.nat({ max: 3 }),
    replacement: fc.constantFrom(...WORDS),
  }),
  fc.record({ kind: fc.constant("delete" as const), block }),
  fc.record({ kind: fc.constant("split" as const), block, at: fc.nat() }),
  fc.record({ kind: fc.constant("merge" as const), block }),
  fc.record({
    kind: fc.constant("insert" as const),
    block,
    text: fc.constantFrom(...WORDS),
  }),
  fc.record({
    kind: fc.constant("align" as const),
    block,
    alignment: fc.constantFrom(...ALIGNMENTS),
  }),
);

const caseArbitrary = fc.record({
  paragraphs: fc.array(
    fc.record({
      words: fc.array(fc.constantFrom(...WORDS), {
        minLength: 2,
        maxLength: 4,
      }),
      alignment: fc.constantFrom(...ALIGNMENTS),
    }),
    { minLength: 2, maxLength: 5 },
  ),
  intents: fc.array(intentArbitrary, { minLength: 1, maxLength: 6 }),
  mode: fc.constantFrom("accept" as const, "reject" as const),
});

/** One intent as an operation on the current blocks after the locked one. */
const operationOf = (
  reviewer: FolioDocxReviewer,
  intent: Intent,
): FolioAIEditOperation | null => {
  const free = reviewer.snapshot().blocks.slice(1);
  const target = free[intent.block % Math.max(free.length, 1)];
  if (target === undefined || target.text.length === 0) return null;
  const id = `op-${intent.kind}`;
  switch (intent.kind) {
    case "replace": {
      const words = target.text.split(" ").filter((word) => word.length > 0);
      const find = words[intent.word % words.length];
      if (find === undefined || words.indexOf(find) !== words.lastIndexOf(find))
        return null;
      return {
        id,
        type: "replaceInBlock",
        blockId: target.id,
        find,
        replace: intent.replacement,
      };
    }
    case "delete":
      return { id, type: "deleteBlock", blockId: target.id };
    case "split": {
      const offset = 1 + (intent.at % Math.max(target.text.length - 1, 1));
      return offset < target.text.length
        ? { id, type: "splitBlock", blockId: target.id, offset }
        : null;
    }
    case "merge":
      return { id, type: "mergeBlockWithNext", blockId: target.id };
    case "insert":
      return {
        id,
        type: "insertAfterBlock",
        blockId: target.id,
        text: intent.text,
      };
    case "align":
      return {
        id,
        type: "setBlockParagraphProperties",
        blockId: target.id,
        properties: { alignment: intent.alignment },
      };
    default: {
      const unhandled: never = intent;
      throw new Error(`Unhandled intent ${JSON.stringify(unhandled)}`);
    }
  }
};

const bodyAfterLockedParagraph = async (
  reviewer: FolioDocxReviewer,
): Promise<string> => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  const body = xml.slice(
    xml.indexOf("<w:body>") + "<w:body>".length,
    xml.indexOf("<w:sectPr"),
  );
  const firstEnd = body.indexOf("</w:p>") + "</w:p>".length;
  // Paragraph and revision ids are allocated as each path goes; the content
  // and properties are what must agree.
  return (
    body
      .slice(firstEnd)
      .replaceAll(/ w14:(?:paraId|textId)="[^"]*"/gu, "")
      .replaceAll(/ w:(?:id|rsid\w*)="[^"]*"/gu, "")
      // `xml:space="preserve"` changes nothing on text with no space at its
      // ends; writers differ on whether they state it there.
      .replaceAll(
        /<w:(t|delText) xml:space="preserve">([^<]*)<\/w:\1>/gu,
        (whole, name: string, text: string) =>
          /^\s|\s$/u.test(text) ? whole : `<w:${name}>${text}</w:${name}>`,
      )
  );
};

describe("resolving around a change a lock holds", () => {
  test("the rest of the story resolves as it does in one pass", async () => {
    const law = (counts: { held: number }) =>
      fc.asyncProperty(caseArbitrary, async ({ paragraphs, intents, mode }) => {
        const source = await createDocx(
          LOCKED_PARAGRAPH +
            paragraphs
              .map(({ words, alignment }) => paragraphXml(words, alignment))
              .join(""),
        );
        const reviewer = await FolioDocxReviewer.fromBuffer(source);
        for (const intent of intents) {
          const operation = operationOf(reviewer, intent);
          if (operation === null) continue;
          reviewer.applyOperations([operation], {
            undefinedReferences: "refuse",
            mode: "tracked-changes",
          });
        }
        const pending = new Uint8Array(await reviewer.toBuffer());

        const locked = await FolioDocxReviewer.fromBuffer(pending);
        const lockedResult =
          mode === "accept" ? locked.acceptAll() : locked.rejectAll();
        expect(lockedResult.skipped.map(({ id }) => id)).toEqual([900]);
        counts.held += 1;

        const unlockedZip = await ZipArchive.read(pending);
        const documentXml =
          (await unlockedZip.entry("word/document.xml")?.text()) ?? "";
        expect(documentXml).toContain(LOCK);
        unlockedZip.set("word/document.xml", documentXml.replace(LOCK, ""));
        const unlocked = await FolioDocxReviewer.fromBuffer(
          await unlockedZip.generate({ compression: "store" }),
        );
        const unlockedResult =
          mode === "accept" ? unlocked.acceptAll() : unlocked.rejectAll();
        expect(unlockedResult.skipped).toEqual([]);

        expect(await bodyAfterLockedParagraph(locked)).toBe(
          await bodyAfterLockedParagraph(unlocked),
        );
        expect(
          locked
            .snapshot()
            .blocks.slice(1)
            .map(({ text }) => text),
        ).toEqual(
          unlocked
            .snapshot()
            .blocks.slice(1)
            .map(({ text }) => text),
        );
      });
    await assertProperty(law({ held: 0 }), {
      id: "resolving around a held change equals one pass",
      numRuns: 150,
    });
    // A fixed sample reaches a held change; each case opens and saves the
    // document several times, and one case is enough.
    const RUNS = 5;
    const counts = { held: 0 };
    await checkCoverageSample(law(counts), RUNS);
    expect(counts.held).toBeGreaterThan(0);
  });
});
