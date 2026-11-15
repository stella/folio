import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { validateDocxPackage } from "@stll/docx-core";
import type { EditorIntent } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import { canonicalSaveParagraphXml } from "../../../../test/canonicalSaveSequence";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { schema } from "../prosemirror/schema";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { canonicalJson } from "../utils/canonicalJson";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DATE = "2026-01-01T00:00:00Z";

/** A range from the start of one body paragraph to the end of another. */
type RangeSpec = {
  kind: "bookmark" | "comment" | "permission";
  from: number;
  to: number;
};
type Fixture = {
  paragraphs: number;
  ranges: readonly RangeSpec[];
  /** Body paragraphs holding a whole complex field each. */
  fields: readonly number[];
  /** A tracked move: its source and its destination paragraph. */
  move: { from: number; to: number } | null;
  /** A table after the first paragraph, with a bookmark that ends after it. */
  table: boolean;
  /** A bookmark whose markers sit between body paragraphs. */
  bodyBookmark: boolean;
};

const paraId = (index: number) =>
  (0x10_00_00_01 + index).toString(16).toUpperCase();

const fieldXml = (index: number) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${String(index + 1)}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const moveXml = (side: "From" | "To") => {
  const range = side === "From" ? 61 : 63;
  return `<w:move${side}RangeStart w:id="${String(range)}" w:name="move1" w:author="A" w:date="${DATE}"/><w:move${side} w:id="${String(range + 1)}" w:author="A" w:date="${DATE}"><w:r>${side === "From" ? "<w:delText>moved</w:delText>" : "<w:t>moved</w:t>"}</w:r></w:move${side}><w:move${side}RangeEnd w:id="${String(range)}"/>`;
};

/**
 * The source body: the last paragraph never holds a marker, so Enter there is
 * an edit beside every range and field the document has.
 */
const documentXml = (fixture: Fixture): string => {
  const opens = new Map<number, string[]>();
  const closes = new Map<number, string[]>();
  const add = (map: Map<number, string[]>, index: number, xml: string) =>
    map.set(index, [...(map.get(index) ?? []), xml]);
  for (const [n, range] of fixture.ranges.entries()) {
    const id = String(n + 1);
    const markers = {
      bookmark: [
        `<w:bookmarkStart w:id="${id}" w:name="_Ref${id}"/>`,
        `<w:bookmarkEnd w:id="${id}"/>`,
      ],
      comment: [
        `<w:commentRangeStart w:id="${id}"/>`,
        `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`,
      ],
      permission: [
        `<w:permStart w:id="${String(40 + n)}" w:edGrp="everyone"/>`,
        `<w:permEnd w:id="${String(40 + n)}"/>`,
      ],
    } as const satisfies Record<RangeSpec["kind"], readonly [string, string]>;
    const [open, close] = markers[range.kind];
    add(opens, range.from, open);
    add(closes, range.to, close);
  }
  for (const index of fixture.fields) add(closes, index, fieldXml(index));
  if (fixture.move) {
    add(closes, fixture.move.from, moveXml("From"));
    add(closes, fixture.move.to, moveXml("To"));
  }
  if (fixture.table) add(closes, 1, '<w:bookmarkEnd w:id="80"/>');
  const paragraph = (index: number) =>
    `<w:p w14:paraId="${paraId(index)}">${(opens.get(index) ?? []).join("")}<w:r><w:t xml:space="preserve">Paragraph ${String(index)} text</w:t></w:r>${(closes.get(index) ?? []).join("")}</w:p>`;
  const blocks: string[] = [];
  for (let index = 0; index < fixture.paragraphs; index++) {
    if (fixture.bodyBookmark && index === 1)
      blocks.push('<w:bookmarkStart w:id="90" w:name="_Body"/>');
    blocks.push(paragraph(index));
    if (fixture.bodyBookmark && index === 1)
      blocks.push('<w:bookmarkEnd w:id="90"/>');
    if (fixture.table && index === 0)
      blocks.push(
        `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p w14:paraId="20000001"><w:bookmarkStart w:id="80" w:name="_Cell"/><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      );
  }
  return `${DECLARATION}\n<w:document xmlns:w="${WORD}" xmlns:w14="${W14}"><w:body>\n${blocks.join("\n")}\n<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
};

const packageBytes = async (fixture: Fixture): Promise<ArrayBuffer> => {
  const comments = fixture.ranges.flatMap((range, n) =>
    range.kind === "comment" ? [String(n + 1)] : [],
  );
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${comments.length > 0 ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : ""}</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${comments.length > 0 ? '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' : ""}</Relationships>`,
  );
  zip.set("word/document.xml", documentXml(fixture));
  if (comments.length > 0)
    zip.set(
      "word/comments.xml",
      `${DECLARATION}<w:comments xmlns:w="${WORD}" xmlns:w14="${W14}">${comments
        .map(
          (id) =>
            `<w:comment w:id="${id}" w:author="Reviewer" w:date="${DATE}" w:initials="R"><w:p w14:paraId="3000000${id}"><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`,
        )
        .join("")}</w:comments>`,
    );
  zip.set(
    "customXml/item1.xml",
    '<opaque xmlns="urn:structural-range-save">  kept &amp; untouched  </opaque>',
  );
  return zip.generate().then((bytes) => bytes.buffer);
};

const fixtureArb: fc.Arbitrary<Fixture> = fc
  .integer({ min: 3, max: 6 })
  .chain((paragraphs) => {
    // Every marker lives before the last paragraph.
    const index = fc.integer({ min: 0, max: paragraphs - 2 });
    return fc.record({
      paragraphs: fc.constant(paragraphs),
      ranges: fc.array(
        fc
          .record({
            kind: fc.constantFrom("bookmark", "comment", "permission"),
            a: index,
            b: index,
          })
          .map(({ kind, a, b }) => ({
            kind,
            from: Math.min(a, b),
            to: Math.max(a, b),
          })),
        { maxLength: 4 },
      ),
      fields: fc.uniqueArray(index, { maxLength: 2 }),
      move: fc.option(
        fc
          .tuple(index, index)
          .filter(([from, to]) => from !== to)
          .map(([from, to]) => ({ from, to })),
        { nil: null },
      ),
      table: fc.boolean(),
      bodyBookmark: fc.boolean(),
    });
  });

type Step =
  | { kind: "enter"; at: number; fraction: number }
  | { kind: "type"; at: number; fraction: number }
  | { kind: "join"; at: number }
  | { kind: "delete"; at: number };

const stepArb: fc.Arbitrary<Step> = fc.oneof(
  fc.record({
    kind: fc.constant("enter" as const),
    at: fc.nat(),
    fraction: fc.double({ min: 0, max: 1, noNaN: true }),
  }),
  fc.record({
    kind: fc.constant("type" as const),
    at: fc.nat(),
    fraction: fc.double({ min: 0, max: 1, noNaN: true }),
  }),
  fc.record({ kind: fc.constant("join" as const), at: fc.nat() }),
  fc.record({ kind: fc.constant("delete" as const), at: fc.nat() }),
);

const bodyParagraphIds = (session: CanonicalSession): string[] =>
  session.document.package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.paraId !== undefined
      ? [block.paraId]
      : [],
  );

/** Commit one intent; a refused edit leaves the session as it was. */
const commit = (
  session: CanonicalSession,
  state: EditorState,
  intent: EditorIntent,
): EditorState | null => {
  const prepared = session.prepareIntent(state, intent);
  if (prepared.isErr()) return null;
  const published = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  });
  return published.isOk() ? published.value.state : null;
};

let freshId = 0x1e_57_00_00;
const intentFor = (
  session: CanonicalSession,
  step: Step,
): EditorIntent | null => {
  const ids = bodyParagraphIds(session);
  const blockId = ids.at(step.at % ids.length);
  if (blockId === undefined) return null;
  const length = session.projection.paragraph(blockId)?.text.length ?? 0;
  switch (step.kind) {
    case "enter":
    case "type": {
      const at = {
        story: "main" as const,
        blockId,
        offset: Math.floor(step.fraction * length),
      };
      if (step.kind === "type")
        return { type: "replaceText", from: at, to: at, text: "x" };
      freshId += 1;
      return {
        type: "splitParagraph",
        at,
        to: at,
        newBlockId: freshId.toString(16).toUpperCase(),
      };
    }
    case "join": {
      const content = session.document.package.document.content;
      const index = content.findIndex(
        (block) => block.type === "paragraph" && block.paraId === blockId,
      );
      const next = content.at(index + 1);
      if (next?.type !== "paragraph" || next.paraId === undefined) return null;
      return {
        type: "joinParagraphs",
        story: "main",
        blockId,
        nextBlockId: next.paraId,
      };
    }
    case "delete":
      return ids.length > 2
        ? { type: "deleteParagraphs", story: "main", blockIds: [blockId] }
        : null;
    default: {
      const unreachable: never = step;
      return unreachable;
    }
  }
};

const partBytes = async (buffer: ArrayBuffer) => {
  const zip = await ZipArchive.read(buffer);
  const parts = new Map<string, Uint8Array>();
  for (const entry of zip.entries())
    // oxlint-disable-next-line no-await-in-loop -- a handful of parts
    parts.set(entry.path, await entry.bytes());
  return parts;
};

/** Each bookmark and permission id opens once and closes once after it. */
const balancedRanges = (xml: string): boolean => {
  for (const kind of ["bookmark", "perm"] as const) {
    const opens = [
      ...xml.matchAll(new RegExp(`<w:${kind}Start\\b[^>]*w:id="(\\d+)"`, "gu")),
    ];
    const closes = [
      ...xml.matchAll(new RegExp(`<w:${kind}End\\b[^>]*w:id="(\\d+)"`, "gu")),
    ];
    const ids = new Set([...opens, ...closes].map((match) => match[1]));
    for (const id of ids) {
      const starts = opens.filter((match) => match[1] === id);
      const ends = closes.filter((match) => match[1] === id);
      if (starts.length !== 1 || ends.length !== 1) return false;
      if ((starts[0]?.index ?? 0) > (ends[0]?.index ?? 0)) return false;
    }
  }
  return true;
};

const documentPart = async (buffer: ArrayBuffer) =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  panic("Expected the main document part");

test(
  "structural saves of documents with bookmarks, fields, comments, moves and permissions splice to the session's document",
  async () => {
    let splicedAfterStructure = 0;
    await assertProperty(
      fc.asyncProperty(
        fixtureArb,
        fc.array(stepArb, { minLength: 1, maxLength: 4 }),
        async (fixture, steps) => {
          const source = await packageBytes(fixture);
          const sourceParts = await partBytes(source);
          const sourceXml = await documentPart(source);
          const opened = await parseDocx(source, { preloadFonts: false });
          const openedComments = canonicalJson(
            opened.package.document.comments ?? [],
          );
          const session = createCanonicalSession(opened).unwrap();
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          // Enter in the one paragraph no range or field touches comes first.
          const plan: Step[] = [
            { kind: "enter", at: fixture.paragraphs - 1, fraction: 1 },
            ...steps,
          ];
          for (const [index, step] of plan.entries()) {
            const intent = intentFor(session, step);
            const next =
              intent === null ? null : commit(session, state, intent);
            if (next === null) {
              if (index === 0) return panic("Enter beside the ranges refused");
              continue;
            }
            state = next;
            const snapshot = session.captureSaveSnapshot();
            const saved = await serializeCanonicalSave({
              snapshot,
              featureFlags: { selectiveSave: true },
            });
            const full = await serializeCanonicalSave({
              snapshot,
              options: { mode: FOLIO_DOCX_SERIALIZATION_MODE.full },
            });
            // A range or field elsewhere in the body never declines the
            // splice of an edit beside it.
            if (index === 0) expect(saved.path).toBe("selective");
            expect(
              await validateDocxPackage(new Uint8Array(saved.buffer)),
            ).toEqual({ valid: true });
            const savedXml = await documentPart(saved.buffer);
            expect(balancedRanges(savedXml)).toBe(true);
            const reopened = await parseDocx(saved.buffer, {
              preloadFonts: false,
            });
            expect(
              describePackageDifferences(snapshot.document, reopened),
            ).toEqual({ messages: [], omitted: 0 });
            expect(
              describePackageDifferences(
                await parseDocx(full.buffer, { preloadFonts: false }),
                reopened,
              ),
            ).toEqual({ messages: [], omitted: 0 });
            if (saved.path !== "selective") continue;
            if (snapshot.structure === "changed") splicedAfterStructure += 1;
            // Every part but the body is the source's, byte for byte (the
            // comments part too while the comments are the opened ones), and
            // so is every source paragraph the session did not touch.
            const commentsKept =
              canonicalJson(reopened.package.document.comments ?? []) ===
              openedComments;
            const savedParts = await partBytes(saved.buffer);
            for (const [path, bytes] of sourceParts)
              if (
                path !== "word/document.xml" &&
                (commentsKept || path !== "word/comments.xml")
              )
                expect([path, savedParts.get(path)]).toEqual([path, bytes]);
            const touched = new Set(
              snapshot.changedBlockIds.map((id) => id.toUpperCase()),
            );
            for (const id of bodyParagraphIds(session))
              if (
                !touched.has(id.toUpperCase()) &&
                sourceXml.includes(`w14:paraId="${id}"`)
              )
                expect(await canonicalSaveParagraphXml(saved.buffer, id)).toBe(
                  await canonicalSaveParagraphXml(source, id),
                );
          }
        },
      ),
      {
        id: "structural saves of documents with bookmarks, fields, comments, moves and permissions splice to the session's document",
        numRuns: 25,
        examples: [
          [
            {
              paragraphs: 4,
              ranges: [
                { kind: "bookmark", from: 0, to: 2 },
                { kind: "comment", from: 1, to: 2 },
                { kind: "permission", from: 0, to: 1 },
              ],
              fields: [1],
              move: { from: 0, to: 2 },
              table: true,
              bodyBookmark: true,
            },
            [
              { kind: "enter", at: 1, fraction: 0.5 },
              { kind: "join", at: 0 },
              { kind: "delete", at: 2 },
            ],
          ],
        ],
      },
    );
    // The generated runs exercised the splice after structural edits.
    expect(splicedAfterStructure).toBeGreaterThan(0);
  },
  propertyTestTimeout(120_000),
);

test("a canonical Enter splices with a bookmark and a field elsewhere in the body", async () => {
  const fixture: Fixture = {
    paragraphs: 3,
    ranges: [{ kind: "bookmark", from: 0, to: 1 }],
    fields: [0],
    move: null,
    table: false,
    bodyBookmark: false,
  };
  const source = await packageBytes(fixture);
  const session = createCanonicalSession(
    await parseDocx(source, { preloadFonts: false }),
  ).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  const intent = intentFor(session, { kind: "enter", at: 2, fraction: 1 });
  if (intent === null || commit(session, state, intent) === null)
    return panic("Enter refused");
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
    featureFlags: { selectiveSave: true },
  });
  expect(saved.path).toBe("selective");
  for (const id of [paraId(0), paraId(1)])
    expect(await canonicalSaveParagraphXml(saved.buffer, id)).toBe(
      await canonicalSaveParagraphXml(source, id),
    );
});
