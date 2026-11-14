/**
 * A public replaceBlock that names a paragraph style restyles the paragraph
 * in the canonical session as the default engine does, direct and tracked,
 * with or without new words, on a fresh package and on one whose joined
 * paragraph went through an accepted or rejected review and a save: both
 * engines give the same package pending, accepted and rejected.
 */

import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../ai-edits";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";

setDefaultTimeout(propertyTestTimeout(240_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const STYLES = `${HEAD}<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>`;

const STYLE_IDS = ["Normal", "Heading1", "Heading2"] as const;

type SourceParagraph = {
  text: string;
  styleId: (typeof STYLE_IDS)[number] | undefined;
  italic: boolean;
  /** Whether the paragraph's mark is pending deletion (a pending join). */
  markDeleted: boolean;
};

const paragraphXml = (
  index: number,
  { text, styleId, italic, markDeleted }: SourceParagraph,
) => {
  const pPr =
    styleId === undefined && !markDeleted
      ? ""
      : `<w:pPr>${styleId === undefined ? "" : `<w:pStyle w:val="${styleId}"/>`}${markDeleted ? `<w:rPr><w:del w:id="${index + 1}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr>` : ""}</w:pPr>`;
  const run =
    text.length === 0
      ? ""
      : `<w:r>${italic ? "<w:rPr><w:i/></w:rPr>" : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  return `<w:p w14:paraId="${(0x10000001 + index).toString(16).toUpperCase()}">${pPr}${run}</w:p>`;
};

const packageOf = async (paragraphs: readonly SourceParagraph[]) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set("word/styles.xml", STYLES);
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs.map((paragraph, index) => paragraphXml(index, paragraph)).join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const WORDS = ["Alpha", "Beta", "Gamma", "Delta beta", ""] as const;

const sourceParagraphArbitrary = fc.record({
  text: fc.constantFrom(...WORDS),
  styleId: fc.constantFrom(...STYLE_IDS, undefined),
  italic: fc.boolean(),
  markDeleted: fc.constant(false),
});

/**
 * Two to four paragraphs; the first one's mark is pending deletion when the
 * package goes through a review first, so that review joins it or keeps it.
 */
const caseArbitrary = fc.record({
  paragraphs: fc.array(sourceParagraphArbitrary, {
    minLength: 2,
    maxLength: 4,
  }),
  review: fc.constantFrom("none", "accept", "reject"),
  target: fc.nat(),
  keepWords: fc.boolean(),
  words: fc.constantFrom(...WORDS.filter((words) => words.length > 0)),
  styleId: fc.constantFrom<string | null | undefined>(
    ...STYLE_IDS,
    null,
    undefined,
  ),
  mode: fc.constantFrom("direct", "tracked-changes"),
});

type Mode = "direct" | "tracked-changes";

/** Each body paragraph's words, style and the italic state of its runs. */
const paragraphsOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            styleId: block.formatting?.styleId,
            runs: block.content.flatMap((item) =>
              item.type === "run"
                ? [
                    {
                      text: item.content
                        .flatMap((content) =>
                          content.type === "text" ? [content.text] : [],
                        )
                        .join(""),
                      italic: item.formatting?.italic === true,
                      bold: item.formatting?.bold === true,
                    },
                  ]
                : [],
            ),
          },
        ]
      : [],
  );
};

/** The package pending, with every change accepted, and with every change rejected. */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return paragraphsOf(await reviewer.toBuffer());
  };
  return {
    pending: await paragraphsOf(bytes),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

const viaDefaultEngine = async (
  bytes: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  const result = reviewer.applyOperations([operation], { mode });
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    bytes: await reviewer.toBuffer(),
  };
};

const viaCanonicalEngine = async (
  bytes: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode, operations: [operation] },
  });
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    refusals: result.issues.flatMap(({ canonicalRefusal }) =>
      canonicalRefusal === undefined ? [] : [canonicalRefusal],
    ),
    bytes: (
      await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
    ).buffer,
  };
};

/** The package the case starts from: fresh, or after its review and a save. */
const startingPackage = async ({
  paragraphs,
  review,
}: {
  paragraphs: readonly SourceParagraph[];
  review: "none" | "accept" | "reject";
}) => {
  if (review === "none") return packageOf(paragraphs);
  const [first, ...rest] = paragraphs;
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await packageOf([
      { ...(first ?? panic("A case has paragraphs.")), markDeleted: true },
      ...rest,
    ]),
  );
  if (review === "accept") reviewer.acceptAll();
  else reviewer.rejectAll();
  return reviewer.toBuffer();
};

type Case = typeof caseArbitrary extends fc.Arbitrary<infer T> ? T : never;

/**
 * One case through both engines, compared. `forCanonical` is what the
 * canonical session is given in place of the operation: the operation
 * itself, or, to show the comparison notices it, one that lost its style.
 */
const compareEngines = async (
  example: Case,
  forCanonical: (operation: FolioDocumentOperation) => FolioDocumentOperation,
) => {
  const bytes = await startingPackage(example);
  const blocks = (
    await FolioDocxReviewer.fromBuffer(bytes.slice(0))
  ).getContent();
  const target =
    blocks[example.target % blocks.length] ??
    panic("The package lost its paragraphs.");
  const operation: FolioDocumentOperation = {
    id: "op",
    type: "replaceBlock",
    blockId: target.id,
    text: example.keepWords ? target.text : example.words,
    ...(example.styleId === undefined ? {} : { styleId: example.styleId }),
  };
  const legacy = await viaDefaultEngine(bytes, operation, example.mode);
  const canonical = await viaCanonicalEngine(
    bytes,
    forCanonical(operation),
    example.mode,
  );
  expect(canonical.refusals).toEqual([]);
  // A blank paragraph given no words and no other style is left as it is:
  // the default engine names that unsupportedBlock, the canonical session a
  // no-op.
  const blankLeftBlank =
    target.text === "" &&
    operation.text === "" &&
    legacy.skipped.join() === "unsupportedBlock";
  expect(canonical.skipped).toEqual(
    blankLeftBlank ? ["noopOperation"] : legacy.skipped,
  );
  expect(await outcomes(canonical.bytes)).toEqual(await outcomes(legacy.bytes));
};

describe("a replacement that names a paragraph style", () => {
  test("restyles the paragraph in both engines alike, pending, accepted and rejected", async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, (example) =>
        compareEngines(example, (operation) => operation),
      ),
      {
        id: "restyles the paragraph in both engines alike, pending, accepted and rejected",
        numRuns: 80,
      },
    );
  });

  test("the comparison notices a canonical replacement that loses its style", async () => {
    const withoutStyle = (
      operation: FolioDocumentOperation,
    ): FolioDocumentOperation =>
      operation.type === "replaceBlock"
        ? {
            id: operation.id,
            type: "replaceBlock",
            blockId: operation.blockId,
            text: operation.text,
          }
        : operation;
    const outcome = await fc.check(
      fc.asyncProperty(caseArbitrary, (example) =>
        compareEngines(example, withoutStyle),
      ),
      propertyConfig({ numRuns: 80 }),
    );
    expect(outcome.failed).toBe(true);
  });

  test("restyles the paragraph a saved accepted or rejected join leaves, with the same words", async () => {
    for (const review of ["accept", "reject"] as const)
      for (const mode of ["direct", "tracked-changes"] as const) {
        const bytes = await startingPackage({
          paragraphs: [
            {
              text: "Alpha",
              styleId: undefined,
              italic: false,
              markDeleted: false,
            },
            {
              text: "Beta",
              styleId: undefined,
              italic: false,
              markDeleted: false,
            },
            {
              text: "Gamma",
              styleId: undefined,
              italic: false,
              markDeleted: false,
            },
          ],
          review,
        });
        const target =
          (await FolioDocxReviewer.fromBuffer(bytes.slice(0)))
            .getContent()
            .find(({ text }) => text.startsWith("Alpha")) ??
          panic("The reviewed package lost the paragraph holding Alpha.");
        const operation: FolioDocumentOperation = {
          id: "op",
          type: "replaceBlock",
          blockId: target.id,
          text: target.text,
          styleId: "Heading1",
        };
        const legacy = await viaDefaultEngine(bytes, operation, mode);
        const canonical = await viaCanonicalEngine(bytes, operation, mode);
        expect(legacy.skipped).toEqual([]);
        expect(canonical.refusals).toEqual([]);
        expect(canonical.skipped).toEqual([]);
        const restyled = await outcomes(canonical.bytes);
        expect(restyled).toEqual(await outcomes(legacy.bytes));
        expect(
          restyled.accepted.find(({ runs }) =>
            runs.some(({ text }) => text.startsWith("Alpha")),
          )?.styleId,
        ).toBe("Heading1");
      }
  });
});
