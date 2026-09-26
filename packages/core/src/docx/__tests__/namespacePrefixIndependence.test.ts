/**
 * Every string-level XML patcher gives the same answer whatever prefix a part
 * binds WordprocessingML to — or refuses.
 *
 * The parser resolves names by namespace URI, so a package spelled `<x:p>`
 * or `<p xmlns="…/main">` opens exactly like the conventional `<w:p>`. The patchers that
 * splice part XML as text (`ensureParaIds`, the selective-save splices, the
 * note-part and numbering patches, the styles append, the comment-range
 * guard every splice goes through) find elements by literal tags, and every
 * fixture folio owns uses the conventional prefixes, so a patcher that
 * only knew `w:` passed every test while reporting success on a respelled
 * package it had not touched. This suite respells each corpus fixture (see
 * `namespacePrefixVariants.ts`) and holds each patcher to: the same result,
 * compared by namespace URI and local name, or an explicit refusal.
 */

import { afterAll, describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../../document-operations";
import type { Document, Paragraph } from "../../types/document";
import { unbalancedCommentRangeIds } from "../commentRangeIntegrity";
import { ensureParaIds, EnsureParaIdsError } from "../ensureParaIds";
import { parseNumbering } from "../numberingParser";
import { parseDocx } from "../parser";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import { createDocx, repackDocx } from "../rezip";
import {
  buildPatchedDocumentXml,
  collectParaIds,
  patchNumberingDefinitions,
} from "../selectiveXmlPatch";
import { serializeDocument } from "../serializer/documentSerializer";
import { serializeNumberingXml } from "../serializer/numberingSerializer";
import { buildStructuralDocumentPatch } from "../structuralXmlPatch";
import {
  canonicalPartTree,
  PREFIX_VARIANTS,
  rewritePackagePrefixes,
  type PrefixVariant,
} from "./namespacePrefixVariants";

const FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
/**
 * Packages folio writes itself, for the parts the corpus lacks: no corpus
 * fixture carries a numbering part.
 */
const GENERATED: Record<string, string> = {
  "generated:lists": "Intro.\n\n- First bullet\n- Second bullet\n\n1. One\n2. Two\n\nOutro.",
};

const FIXTURES = [
  ...readdirSync(FIXTURES_DIR)
    .filter((name) => name.endsWith(".docx"))
    .sort(),
  ...Object.keys(GENERATED),
];

const loadFixture = async (fixture: string): Promise<Uint8Array> => {
  const markdown = GENERATED[fixture];
  return markdown === undefined
    ? new Uint8Array(readFileSync(path.join(FIXTURES_DIR, fixture)))
    : new Uint8Array(await createDocx(fromMarkdown(markdown)));
};

const TARGET_PART = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/iu;

const toArrayBuffer = (bytes: Uint8Array | ArrayBuffer): ArrayBuffer =>
  bytes instanceof ArrayBuffer
    ? bytes
    : (bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);

const readParts = async (docx: Uint8Array | ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await JSZip.loadAsync(docx);
  const parts = new Map<string, string>();
  for (const [name, file] of Object.entries(zip.files)) {
    const lower = name.toLowerCase();
    if (
      !file.dir &&
      lower.startsWith("word/") &&
      lower.endsWith(".xml") &&
      !lower.includes("_rels")
    ) {
      // oxlint-disable-next-line no-await-in-loop -- a handful of parts per fixture
      parts.set(name, await file.async("text"));
    }
  }
  return parts;
};

/** Drop every paragraph id folio's own scan would recognize, so `ensureParaIds` has work. */
const withoutParagraphIds = async (docx: Uint8Array): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(docx);
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir || !TARGET_PART.test(name)) continue;
    // oxlint-disable-next-line no-await-in-loop -- a handful of parts per fixture
    const xml = await file.async("text");
    zip.file(name, xml.replace(/\sw14:(?:paraId|textId)=(["'])[^"']*\1/gu, ""));
  }
  return zip.generateAsync({ type: "uint8array" });
};

/** The two sides of one comparison: the fixture as written, and respelled. */
type Pair = { canonical: Uint8Array; variant: Uint8Array };

const respell = async (canonical: Uint8Array, variant: PrefixVariant): Promise<Pair | null> => {
  const respelled = await rewritePackagePrefixes(canonical, variant);
  return respelled === null ? null : { canonical, variant: respelled.docx };
};

/** Both packages, part by part, compared by meaning rather than spelling. */
const expectSameParts = async (
  canonical: Uint8Array | ArrayBuffer,
  variant: Uint8Array | ArrayBuffer,
): Promise<void> => {
  const [canonicalParts, variantParts] = await Promise.all([
    readParts(canonical),
    readParts(variant),
  ]);
  expect([...variantParts.keys()].sort()).toEqual([...canonicalParts.keys()].sort());
  for (const [name, xml] of canonicalParts) {
    const other = variantParts.get(name);
    if (other === undefined) continue;
    expect({ name, tree: canonicalPartTree(other) }).toEqual({
      name,
      tree: canonicalPartTree(xml),
    });
  }
};

const blockRows = async (docx: ArrayBuffer) =>
  (await FolioDocxReviewer.fromBuffer(docx)).getContent().map((block) => ({
    id: block.id,
    stability: block.idStability ?? "package",
    text: block.text,
  }));

const firstTextParagraph = (doc: Document): Paragraph | undefined =>
  doc.package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" &&
      block.paraId !== undefined &&
      block.content.some(
        (item) => item.type === "run" && item.content.some((run) => run.type === "text"),
      ),
  );

/** Append to the paragraph's first text node; the edit every save path below makes. */
const editParagraph = (paragraph: Paragraph): void => {
  for (const item of paragraph.content) {
    if (item.type !== "run") continue;
    for (const run of item.content) {
      if (run.type === "text") {
        run.text = `${run.text} [edited]`;
        return;
      }
    }
  }
};

const cases = FIXTURES.flatMap((fixture) =>
  PREFIX_VARIANTS.map((variant) => ({ fixture, variant })),
);

/**
 * How often each check compared two real results, per variant. A respelling
 * or an edit that silently stopped applying would otherwise leave the suite
 * green by comparing nothing.
 */
const exercised = new Map<string, number>();
const count = (check: string, variant: PrefixVariant): void => {
  const key = `${variant}:${check}`;
  exercised.set(key, (exercised.get(key) ?? 0) + 1);
};

/** A patcher's answer on the respelled part: null (refused), or the canonical answer. */
const sameOrRefused = (
  check: string,
  variant: PrefixVariant,
  canonical: string | null,
  respelled: string | null,
): void => {
  if (canonical !== null) {
    count(check, variant);
  }
  if (respelled === null) {
    return;
  }
  expect(canonical).not.toBeNull();
  expect(canonicalPartTree(respelled)).toBe(canonicalPartTree(canonical ?? ""));
};

describe("string-level XML patchers are prefix independent", () => {
  afterAll(() => {
    for (const variant of PREFIX_VARIANTS) {
      for (const check of ["ensureParaIds", "repack", "review"]) {
        expect({
          check: `${variant}:${check}`,
          atLeast: (exercised.get(`${variant}:${check}`) ?? 0) >= 20,
        }).toEqual({
          check: `${variant}:${check}`,
          atLeast: true,
        });
      }
      for (const check of ["paragraphSplice", "numbering"]) {
        expect({
          check: `${variant}:${check}`,
          any: (exercised.get(`${variant}:${check}`) ?? 0) > 0,
        }).toEqual({
          check: `${variant}:${check}`,
          any: true,
        });
      }
    }
  });

  test.each(cases)("$fixture as $variant", async ({ fixture, variant }) => {
    const original = await loadFixture(fixture);

    // ensureParaIds: same counts and the same stamped parts, or a named refusal.
    const stripped = await respell(await withoutParagraphIds(original), variant);
    if (stripped === null) {
      return;
    }
    const canonicalIds = await ensureParaIds(stripped.canonical);
    let variantIds: Awaited<ReturnType<typeof ensureParaIds>>;
    try {
      variantIds = await ensureParaIds(stripped.variant);
    } catch (error) {
      expect(error).toBeInstanceOf(EnsureParaIdsError);
      return;
    }
    expect({
      assigned: variantIds.assigned,
      deduplicated: variantIds.deduplicated,
      alreadyComplete: variantIds.alreadyComplete,
    }).toEqual({
      assigned: canonicalIds.assigned,
      deduplicated: canonicalIds.deduplicated,
      alreadyComplete: canonicalIds.alreadyComplete,
    });
    await expectSameParts(canonicalIds.docx, variantIds.docx);
    count("ensureParaIds", variant);

    // From here on both sides carry the same ids: the canonical normalized
    // package, and that package respelled.
    const pair = await respell(new Uint8Array(canonicalIds.docx), variant);
    if (pair === null) {
      return;
    }
    const [canonicalParts, variantParts] = await Promise.all([
      readParts(pair.canonical),
      readParts(pair.variant),
    ]);
    const partOf = (parts: Map<string, string>, lower: string) =>
      [...parts].find(([name]) => name.toLowerCase() === lower)?.[1];

    // The comment-range guard every splice goes through sees the same ranges.
    for (const [name, xml] of canonicalParts) {
      expect([...unbalancedCommentRangeIds(variantParts.get(name) ?? "")].sort()).toEqual(
        [...unbalancedCommentRangeIds(xml)].sort(),
      );
    }

    const doc = await parseDocx(toArrayBuffer(pair.canonical));
    const paragraph = firstTextParagraph(doc);
    const canonicalDocXml = partOf(canonicalParts, "word/document.xml") ?? "";
    const variantDocXml = partOf(variantParts, "word/document.xml") ?? "";
    if (paragraph?.paraId !== undefined) {
      editParagraph(paragraph);
      const serialized = serializeDocument(doc);
      const changed = new Set([paragraph.paraId]);
      if (collectParaIds(serialized).has(paragraph.paraId)) {
        sameOrRefused(
          "paragraphSplice",
          variant,
          buildPatchedDocumentXml(canonicalDocXml, serialized, changed),
          buildPatchedDocumentXml(variantDocXml, serialized, changed),
        );
        const structural = (originalXml: string) =>
          buildStructuralDocumentPatch({
            originalXml,
            serializedXml: serialized,
            changedIds: changed,
          });
        sameOrRefused(
          "structuralSplice",
          variant,
          structural(canonicalDocXml),
          structural(variantDocXml),
        );
      }
    }

    // Numbering: an edited level text spliced into the source part.
    const canonicalNumbering = partOf(canonicalParts, "word/numbering.xml");
    const variantNumbering = partOf(variantParts, "word/numbering.xml");
    if (canonicalNumbering !== undefined && variantNumbering !== undefined) {
      const definitions = parseNumbering(canonicalNumbering).definitions;
      const level = definitions.abstractNums[0]?.levels[0];
      if (level) {
        const edited = structuredClone(definitions);
        const editedLevel = edited.abstractNums[0]?.levels[0];
        if (editedLevel) editedLevel.lvlText = `${level.lvlText ?? ""}#`;
        const currentXml = serializeNumberingXml(edited);
        const patch = (xml: string) =>
          patchNumberingDefinitions({
            originalXml: xml,
            baselineXml: serializeNumberingXml(parseNumbering(xml).definitions),
            currentXml,
          });
        sameOrRefused("numbering", variant, patch(canonicalNumbering), patch(variantNumbering));
      }
    }

    // Full repack of the same model edit (body paragraph and first footnote,
    // which runs the note-part patch and the styles append): same package.
    const repacked = async (docx: Uint8Array) => {
      const parsed = await parseDocx(toArrayBuffer(docx));
      const target = firstTextParagraph(parsed);
      if (target) editParagraph(target);
      const note = parsed.package.footnotes?.[0]?.content.find(
        (block): block is Paragraph => block.type === "paragraph",
      );
      if (note) editParagraph(note);
      return repackDocx(parsed);
    };
    const [canonicalRepack, variantRepack] = await Promise.all([
      repacked(pair.canonical),
      repacked(pair.variant),
    ]);
    expect(await blockRows(variantRepack)).toEqual(await blockRows(canonicalRepack));
    await expectSameParts(canonicalRepack, variantRepack);
    count("repack", variant);

    // The reviewer's save, which tries the selective splice first.
    const reviewed = async (docx: Uint8Array) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(toArrayBuffer(docx));
      const block = reviewer.getContent().find((candidate) => candidate.text.trim().length > 0);
      const issues = block
        ? reviewer.applyDocumentOperations({
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "direct",
            operations: [
              {
                id: "edit",
                type: "replaceInBlock",
                blockId: block.id,
                find: block.text,
                replace: `${block.text} [edited]`,
              },
            ],
          }).issues
        : [];
      return { issues, saved: await reviewer.toBuffer() };
    };
    const [canonicalReview, variantReview] = await Promise.all([
      reviewed(pair.canonical),
      reviewed(pair.variant),
    ]);
    expect(variantReview.issues).toEqual(canonicalReview.issues);
    expect(await blockRows(variantReview.saved)).toEqual(await blockRows(canonicalReview.saved));
    count("review", variant);
  });

  test.each(PREFIX_VARIANTS)(
    "a style the model adds is appended to a %s styles part, not rewritten over it",
    async (variant) => {
      const canonical = await loadFixture("generated:lists");
      const respelled = await rewritePackagePrefixes(canonical, variant);
      if (!respelled?.rewritten.includes("word/styles.xml")) {
        throw new Error(`generated fixture has no ${variant} styles part`);
      }
      const addStyle = async (docx: Uint8Array) => {
        const doc = await parseDocx(toArrayBuffer(docx));
        const styles = doc.package.styles?.styles ?? [];
        const base = styles.find((style) => style.type === "paragraph");
        if (!base) throw new Error("no paragraph style to clone");
        styles.push({ ...structuredClone(base), styleId: "AddedByModel", default: false });
        return repackDocx(doc);
      };
      const [canonicalOut, variantOut] = await Promise.all([
        addStyle(canonical),
        addStyle(respelled.docx),
      ]);
      const [sourceParts, variantParts] = await Promise.all([
        readParts(respelled.docx),
        readParts(variantOut),
      ]);
      const styles = variantParts.get("word/styles.xml") ?? "";
      // Appended before the root's own close tag; the source bytes kept.
      const source = sourceParts.get("word/styles.xml") ?? "";
      const close = source.lastIndexOf("</");
      expect(styles.startsWith(source.slice(0, close))).toBe(true);
      expect(styles).toContain('styleId="AddedByModel"');
      await expectSameParts(canonicalOut, variantOut);
    },
  );
});
