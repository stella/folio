import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

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
const RELS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PARAGRAPHS = 4;

const paraId = (index: number) =>
  (0x10_00_00_01 + index).toString(16).toUpperCase();

/** Four body paragraphs, a comment on the first, and a default header. */
const sourcePackage = async (): Promise<ArrayBuffer> => {
  const paragraphs = Array.from({ length: PARAGRAPHS }, (_, index) => {
    const text = `<w:r><w:t xml:space="preserve">Paragraph ${String(index)} text</w:t></w:r>`;
    return index === 0
      ? `<w:p w14:paraId="${paraId(index)}"><w:commentRangeStart w:id="1"/>${text}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`
      : `<w:p w14:paraId="${paraId(index)}">${text}</w:p>`;
  });
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="${RELS}/comments" Target="comments.xml"/><Relationship Id="rId8" Type="${RELS}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${DECLARATION}\n<w:document xmlns:w="${WORD}" xmlns:w14="${W14}" xmlns:r="${RELS}"><w:body>\n${paragraphs.join("\n")}\n<w:sectPr><w:headerReference w:type="default" r:id="rId8"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `${DECLARATION}<w:comments xmlns:w="${WORD}" xmlns:w14="${W14}"><w:comment w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R"><w:p w14:paraId="30000001"><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  zip.set(
    "word/header1.xml",
    `${DECLARATION}<w:hdr xmlns:w="${WORD}" xmlns:w14="${W14}"><w:p w14:paraId="40000001"><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
  );
  return zip.generate().then((bytes) => bytes.buffer);
};

type Step =
  | { kind: "enter" | "type"; at: number; fraction: number }
  | { kind: "join" | "delete"; at: number }
  | { kind: "undo" | "save" };

const stepArb: fc.Arbitrary<Step> = fc.oneof(
  fc.record({
    kind: fc.constantFrom("enter" as const, "type" as const),
    at: fc.nat(),
    fraction: fc.double({ min: 0, max: 1, noNaN: true }),
  }),
  fc.record({
    kind: fc.constantFrom("join" as const, "delete" as const),
    at: fc.nat(),
  }),
  fc.record({ kind: fc.constantFrom("undo" as const, "save" as const) }),
);

const bodyParagraphIds = (session: CanonicalSession): string[] =>
  session.document.package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.paraId !== undefined
      ? [block.paraId]
      : [],
  );

let freshId = 0x1e_57_00_00;
const intentFor = (
  session: CanonicalSession,
  step: Extract<Step, { at: number }>,
): EditorIntent | null => {
  const ids = bodyParagraphIds(session);
  const blockId = ids.at(step.at % ids.length);
  if (blockId === undefined) return null;
  switch (step.kind) {
    case "enter":
    case "type": {
      const length = session.projection.paragraph(blockId)?.text.length ?? 0;
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
      const next = ids.at((step.at % ids.length) + 1);
      return next === undefined
        ? null
        : { type: "joinParagraphs", story: "main", blockId, nextBlockId: next };
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

/** Publish a prepared commit; a refused one leaves the session as it was. */
const publish = (
  session: CanonicalSession,
  state: EditorState,
  prepared: ReturnType<CanonicalSession["prepareIntent"]>,
): EditorState => {
  if (prepared.isErr()) return state;
  const published = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  });
  return published.isOk() ? published.value.state : state;
};

const partBytes = async (buffer: ArrayBuffer) => {
  const parts = new Map<string, Uint8Array>();
  for (const entry of (await ZipArchive.read(buffer)).entries())
    // oxlint-disable-next-line no-await-in-loop -- a handful of parts
    parts.set(entry.path, await entry.bytes());
  return parts;
};
const documentPart = async (buffer: ArrayBuffer) =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  panic("Expected the main document part");

/** Make the save at `version` the baseline, as the editor does once saved. */
const rebase = async (
  session: CanonicalSession,
  version: number,
  buffer: ArrayBuffer,
): Promise<boolean> =>
  session.beginSaveBaseline(version) &&
  session.completeSaveBaseline(
    version,
    await parseDocx(buffer.slice(0), { preloadFonts: false }),
  );

test(
  "saves after a save patch the previous save and reopen to the session's document, across undo",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(stepArb, { minLength: 1, maxLength: 8 }),
        async (steps) => {
          const source = await sourcePackage();
          const session = createCanonicalSession(
            await parseDocx(source, { preloadFonts: false }),
          ).unwrap();
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          let previous = source;
          let previousComments = canonicalJson(
            session.document.package.document.comments ?? [],
          );
          for (const step of [...steps, { kind: "save" } as const]) {
            if (step.kind === "undo") {
              state = publish(session, state, session.prepareUndo(state));
              continue;
            }
            if (step.kind !== "save") {
              const intent = intentFor(session, step);
              if (intent !== null)
                state = publish(
                  session,
                  state,
                  session.prepareIntent(state, intent),
                );
              continue;
            }
            const snapshot = session.captureSaveSnapshot();
            const saved = await serializeCanonicalSave({
              snapshot,
              featureFlags: { selectiveSave: true },
            });
            const full = await serializeCanonicalSave({
              snapshot,
              options: { mode: FOLIO_DOCX_SERIALIZATION_MODE.full },
            });
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
            const comments = canonicalJson(
              reopened.package.document.comments ?? [],
            );
            if (saved.path === "selective") {
              // Every part but the body is the previous save's, byte for
              // byte (the comments part while the comments are the same),
              // and so is every paragraph untouched since that save.
              const before = await partBytes(previous);
              const after = await partBytes(saved.buffer);
              for (const [path, bytes] of before)
                if (
                  path !== "word/document.xml" &&
                  (path !== "word/comments.xml" ||
                    comments === previousComments)
                )
                  expect([path, after.get(path)]).toEqual([path, bytes]);
              const touched = new Set(
                snapshot.changedBlockIds.map((id) => id.toUpperCase()),
              );
              const previousXml = await documentPart(previous);
              for (const id of bodyParagraphIds(session))
                if (
                  !touched.has(id.toUpperCase()) &&
                  previousXml.includes(`w14:paraId="${id}"`)
                )
                  expect(
                    await canonicalSaveParagraphXml(saved.buffer, id),
                  ).toBe(await canonicalSaveParagraphXml(previous, id));
            }
            expect(await rebase(session, snapshot.version, saved.buffer)).toBe(
              true,
            );
            // Nothing has changed since the baseline.
            const rebased = session.captureSaveSnapshot();
            expect([rebased.changedBlockIds, rebased.structure]).toEqual([
              [],
              "stable",
            ]);
            previous = saved.buffer;
            previousComments = comments;
          }
        },
      ),
      {
        id: "saves after a save patch the previous save and reopen to the session's document, across undo",
        numRuns: 25,
        examples: [
          [
            [
              { kind: "enter", at: 1, fraction: 0.5 },
              { kind: "save" },
              { kind: "type", at: 3, fraction: 0 },
              { kind: "save" },
              { kind: "undo" },
              { kind: "undo" },
              { kind: "save" },
              { kind: "delete", at: 0 },
            ],
          ],
        ],
      },
    );
  },
  propertyTestTimeout(120_000),
);

test("a text edit after a structural save takes the text patch again", async () => {
  const session = createCanonicalSession(
    await parseDocx(await sourcePackage(), { preloadFonts: false }),
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const enter = intentFor(session, { kind: "enter", at: 2, fraction: 1 });
  if (enter === null) return panic("Expected Enter");
  state = publish(session, state, session.prepareIntent(state, enter));
  const first = session.captureSaveSnapshot();
  expect(first.structure).toBe("changed");
  const saved = await serializeCanonicalSave({
    snapshot: first,
    featureFlags: { selectiveSave: true },
  });
  expect(await rebase(session, first.version, saved.buffer)).toBe(true);
  const type = intentFor(session, { kind: "type", at: 0, fraction: 0 });
  if (type === null) return panic("Expected typing");
  publish(session, state, session.prepareIntent(state, type));
  const second = session.captureSaveSnapshot();
  expect([second.structure, second.changedBlockIds]).toEqual([
    "stable",
    [paraId(0)],
  ]);
});

test("a save baseline releases the one before it", async () => {
  const session = createCanonicalSession(
    await parseDocx(await sourcePackage(), { preloadFonts: false }),
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const saveOnce = async (at: number): Promise<WeakRef<ArrayBuffer>> => {
    const intent = intentFor(session, { kind: "type", at, fraction: 0 });
    if (intent === null) return panic("Expected typing");
    state = publish(session, state, session.prepareIntent(state, intent));
    const snapshot = session.captureSaveSnapshot();
    const saved = await serializeCanonicalSave({
      snapshot,
      featureFlags: { selectiveSave: true },
    });
    const reopened = await parseDocx(saved.buffer.slice(0), {
      preloadFonts: false,
    });
    const buffer = reopened.originalBuffer ?? panic("Expected a baseline");
    expect(session.beginSaveBaseline(snapshot.version)).toBe(true);
    expect(session.completeSaveBaseline(snapshot.version, reopened)).toBe(true);
    return new WeakRef(buffer);
  };
  const baselines: WeakRef<ArrayBuffer>[] = [];
  for (let round = 0; round < 4; round++)
    // oxlint-disable-next-line no-await-in-loop -- each save rebases the next
    baselines.push(await saveOnce(round));
  await new Promise((resolve) => setTimeout(resolve, 0));
  Bun.gc(true);
  await new Promise((resolve) => setTimeout(resolve, 0));
  Bun.gc(true);
  // Only the latest baseline is still held.
  expect(baselines.map((baseline) => baseline.deref() !== undefined)).toEqual([
    false,
    false,
    false,
    true,
  ]);
});
