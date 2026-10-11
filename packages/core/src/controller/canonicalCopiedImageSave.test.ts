import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import path from "node:path";
import { panic, Result } from "better-result";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { DOCUMENT_OP_TYPES, OP_STORIES, storyBody, type OpStory } from "@stll/docx-core/ops";
import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { canonicalJson } from "../utils/canonicalJson";
import { normalizeDrawingIds } from "../docx/drawingIdNormalization";
import { imageDerivedFactsOf, resolveImageData } from "../docx/imageParser";
import { parseDocx } from "../docx/parser";
import { CanonicalImageRelationshipRefusalError, createDocx } from "../docx/rezip";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { createSimpleDocument } from "../docx/serializer/documentSerializer";
import { visitParagraphRuns } from "../docx/paragraphTraversal";
import { parseRelationships, RELATIONSHIP_TYPES, resolveRelativePath } from "../docx/relsParser";
import {
  findAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseXmlDocument,
  type XmlElement,
} from "../docx/xmlParser";
import { schema } from "../prosemirror/schema";
import type { Document, DrawingContent, Image, Paragraph } from "../types/document";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { canonicalSaveParagraphXml } from "../../../../test/canonicalSaveSequence";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { createCanonicalSession } from "./canonicalSession";
import { prepareCanonicalPaste } from "./canonicalClipboard";

setDefaultTimeout(propertyTestTimeout(30_000));
const STORIES = {
  main: OP_STORIES.MAIN,
  header: { kind: "header", rId: "rIdCopiedHeader" },
  footer: { kind: "footer", rId: "rIdCopiedFooter" },
  footnote: { kind: "footnote", id: 31 },
  endnote: { kind: "endnote", id: 32 },
} as const satisfies Record<"main" | Exclude<OpStory, "main">["kind"], OpStory>;
const MODES = ["editing", "suggesting"] as const;
const RESOURCES = ["package", "dataUrl", "http"] as const;
const CASES = Object.entries(STORIES).flatMap(([kind, story]) =>
  MODES.flatMap((mode) => RESOURCES.map((resource) => ({ kind, story, mode, resource }))),
);
const MAIN_ID = "76000001";
const STORY_ID = "76000003";
const EXTERNAL_URL = "https://example.com/copied.png";
const DRAWINGML_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);
const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: "Story" }] }],
});
const drawings = (document: Document, story: OpStory): DrawingContent[] => {
  const result: DrawingContent[] = [];
  for (const block of storyBody(document, story).content) {
    if (block.type !== "paragraph") continue;
    visitParagraphRuns(block, (run) => {
      for (const item of run.content) if (item.type === "drawing") result.push(item);
    });
  }
  return result;
};
const relationshipPath = (part: string) =>
  `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`;
const owningPart = async (zip: JSZip, story: OpStory): Promise<string> => {
  if (story === OP_STORIES.MAIN) return "word/document.xml";
  if (story.kind === "footnote") return "word/footnotes.xml";
  if (story.kind === "endnote") return "word/endnotes.xml";
  const xml = await zip.file("word/_rels/document.xml.rels")?.async("text");
  const relation = parseRelationships(xml ?? "").get(story.rId);
  if (!relation) panic("Missing serialized story binding");
  return resolveRelativePath("word/_rels/document.xml.rels", relation.target);
};
const imageBindings = (xml: string): string[] => {
  const ids: string[] = [];
  const visit = (element: XmlElement) => {
    if (
      getLocalName(element.name ?? "") === "blip" &&
      DRAWINGML_URIS.has(getNamespaceUri(element) ?? "")
    ) {
      for (const field of ["embed", "link"]) {
        const attribute = findAttributeByNamespaceUri(
          element,
          OFFICE_RELATIONSHIP_NAMESPACE_URIS,
          field,
        );
        if (attribute) ids.push(attribute.value);
      }
    }
    for (const child of getChildElements(element)) visit(child);
  };
  const root = parseXmlDocument(xml);
  if (!root) panic("Saved owner XML is invalid");
  visit(root);
  return ids;
};

type DestinationOptions = { story: OpStory; conflictingBinding?: "present" };
const destination = async ({ story, conflictingBinding }: DestinationOptions) => {
  const model = createSimpleDocument([{ text: "Target" }, { text: "Untouched" }]);
  for (const [index, block] of model.package.document.content.entries())
    if (block.type === "paragraph") block.paraId = index === 0 ? MAIN_ID : "76000002";
  if (story !== OP_STORIES.MAIN) {
    const content = [paragraph(STORY_ID)];
    switch (story.kind) {
      case "header":
        model.package.headers = new Map([
          [story.rId, { type: "header", hdrFtrType: "default", content }],
        ]);
        model.package.document.finalSectionProperties = {
          headerReferences: [{ type: "default", rId: story.rId }],
        };
        break;
      case "footer":
        model.package.footers = new Map([
          [story.rId, { type: "footer", hdrFtrType: "default", content }],
        ]);
        model.package.document.finalSectionProperties = {
          footerReferences: [{ type: "default", rId: story.rId }],
        };
        break;
      case "footnote":
      case "endnote": {
        const first = model.package.document.content.at(0);
        if (first?.type !== "paragraph") panic("Missing destination paragraph");
        if (story.kind === "footnote") {
          model.package.footnotes = [{ type: "footnote", id: story.id, content }];
          first.content.push({ type: "run", content: [{ type: "footnoteRef", id: story.id }] });
        } else {
          model.package.endnotes = [{ type: "endnote", id: story.id, content }];
          first.content.push({ type: "run", content: [{ type: "endnoteRef", id: story.id }] });
        }
        break;
      }
    }
  }
  const zip = await JSZip.loadAsync(await createDocx(model));
  let conflictId: string | undefined;
  if (conflictingBinding) {
    const xml = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const rootRelationships = parseRelationships(xml ?? "");
    let identity = 1;
    while (rootRelationships.has(`rId${identity}`)) identity++;
    conflictId = `rId${identity}`;
    const part = await owningPart(zip, story);
    zip.file(
      relationshipPath(part),
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${conflictId}" Type="${RELATIONSHIP_TYPES.image}" Target="https://example.com/original.png" TargetMode="External"/></Relationships>`,
    );
  }
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  return { document: await parseDocx(buffer, { preloadFonts: false }), conflictId };
};

type CopyOptions = {
  source: Document;
  story: OpStory;
  mode: (typeof MODES)[number];
  resource: (typeof RESOURCES)[number];
  offset: number;
};
const copyImage = async ({ source, story, mode, resource, offset }: CopyOptions) => {
  const imageDocument = await parseDocx(await shapeArrayBuffer("image"), { preloadFonts: false });
  const projectedImage = createCanonicalSession(imageDocument).unwrap().projection.doc;
  let imageNode: PMNode | undefined;
  projectedImage.descendants((node) => {
    if (node.type.name === "image") imageNode = node;
  });
  if (!imageNode) panic("Package fixture has no projected picture");
  const session = createCanonicalSession(source).unwrap();
  session.setMode(
    mode === "editing" ? { type: "editing" } : { type: "suggesting", author: "Image author" },
  );
  let bodyState = EditorState.create({ schema, doc: session.projection.doc });
  bodyState = bodyState.apply(
    bodyState.tr.setSelection(TextSelection.create(bodyState.doc, 1 + (offset % 7))),
  );
  const before = session.captureSaveSnapshot().document;
  const node = imageNode.type.create({
    ...imageNode.attrs,
    ...(resource === "package"
      ? {}
      : {
          rId: null,
          src:
            resource === "http"
              ? EXTERNAL_URL
              : drawings(imageDocument, OP_STORIES.MAIN).at(0)?.image.src,
        }),
  });
  const paste = prepareCanonicalPaste({
    session,
    state: bodyState,
    slice: new Slice(Fragment.from(node), 0, 0),
    ...(resource === "package" ? { sourceDocument: imageDocument } : {}),
  }).unwrap();
  bodyState = bodyState.apply(paste.transaction);
  paste.publish().unwrap();
  const allocated =
    drawings(session.document, OP_STORIES.MAIN).at(0) ?? panic("Clipboard allocated no picture");
  expect(allocated.image.rId).toBeDefined();
  expect(allocated.image.rId).not.toBe(drawings(imageDocument, OP_STORIES.MAIN).at(0)?.image.rId);
  let storyState = bodyState;
  if (story !== OP_STORIES.MAIN) {
    session.breakUndoGroup();
    const projection = session.projectStory(story).unwrap();
    storyState = EditorState.create({ schema, doc: projection.doc });
    const insert = session
      .prepareOps(
        storyState,
        [
          {
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at: { story, blockId: STORY_ID, offset: offset % 6 },
            slice: {
              content: [{ type: "run", content: [structuredClone(allocated)] }],
              openStart: 0,
              openEnd: 0,
            },
            ...(mode === "suggesting"
              ? { revision: { id: 900001, author: "Image author", date: "2026-10-05T12:00:00Z" } }
              : {}),
          },
        ],
        projection.selectionAt(storyState).unwrap(),
      )
      .unwrap();
    storyState = storyState.apply(insert.transaction);
    insert.publish().unwrap();
  }
  return { session, before, bodyState, storyState, allocated };
};

const assertSavedOwner = async (zip: JSZip, model: Document, story: OpStory) => {
  const part = await owningPart(zip, story);
  const relsPath = relationshipPath(part);
  const local = parseRelationships((await zip.file(relsPath)?.async("text")) ?? "");
  const bindings = imageBindings((await zip.file(part)?.async("text")) ?? "");
  const pictures = drawings(model, story);
  expect(pictures).toHaveLength(1);
  for (const { image } of pictures) {
    const id = image.rId ?? panic("Picture lost allocated identity");
    const desired = model.package.relationships?.get(id) ?? panic("Allocated relationship missing");
    const saved = local.get(id) ?? panic("Saved owner lost allocated relationship");
    expect(bindings).toContain(id);
    expect(saved.type).toBe(RELATIONSHIP_TYPES.image);
    expect(saved.targetMode).toBe(desired.targetMode);
    if (desired.targetMode === "External") {
      expect(saved.target).toBe(desired.target);
    } else {
      const expectedPath = resolveRelativePath("word/_rels/document.xml.rels", desired.target);
      expect(resolveRelativePath(relsPath, saved.target)).toBe(expectedPath);
      const media = model.package.media?.get(expectedPath) ?? panic("Allocated bytes missing");
      expect(await zip.file(expectedPath)?.async("uint8array")).toEqual(new Uint8Array(media.data));
    }
  }
};

const hasStory = (document: Document, story: OpStory) => {
  if (story === OP_STORIES.MAIN) return true;
  switch (story.kind) {
    case "header":
      return document.package.headers?.has(story.rId) ?? false;
    case "footer":
      return document.package.footers?.has(story.rId) ?? false;
    case "footnote":
      return document.package.footnotes?.some(({ id }) => id === story.id) ?? false;
    case "endnote":
      return document.package.endnotes?.some(({ id }) => id === story.id) ?? false;
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

/** Project only serializer-generated facts absent from the authored canonical drawing. */
const exportedPackageOf = (document: Document) => {
  const expected = structuredClone(document);
  normalizeDrawingIds({
    documentBody: expected.package.document,
    headers: expected.package.headers,
    footers: expected.package.footers,
    footnotes: expected.package.footnotes,
    endnotes: expected.package.endnotes,
  });
  for (const story of Object.values(STORIES)) {
    if (!hasStory(expected, story)) continue;
    for (const { image } of drawings(expected, story)) {
      const relationship =
        image.rId === undefined ? undefined : expected.package.relationships?.get(image.rId);
      if (!relationship) panic("Expected canonical drawing lacks its allocated relationship");
      // Only the parser's actual derived-field producer can enrich missing fields.
      const facts = imageDerivedFactsOf({
        id: image.id,
        pictureNames: image.pictureNames ?? { name: image.filename ?? `image${image.id}` },
        imageData: resolveImageData(
          image.rId,
          expected.package.relationships,
          expected.package.media,
        ),
      });
      for (const [field, value] of Object.entries(facts))
        if (Reflect.get(image, field) === undefined) Reflect.set(image, field, value);
      if (relationship.targetMode === "External") delete image.src;
    }
  }
  for (const media of [...(expected.package.media?.values() ?? [])]) {
    // Render URLs are lazy, non-enumerable parser caches. Exact ZIP bytes and
    // each reopened drawing's resolved src are checked independently above.
    delete media.dataUrl;
    if (media.filename === undefined) media.filename = media.path.split("/").at(-1);
    if (media.path.startsWith("word/"))
      expected.package.media?.set(media.path.slice("word/".length), media);
  }
  return expected;
};

test.each(CASES)(
  "copied $resource image saves with exact $kind ownership in $mode",
  async ({ story, mode, resource }) => {
    await assertProperty(
      fc.asyncProperty(fc.nat({ max: 12 }), async (offset) => {
        const source = await destination({ story });
        const driver = await copyImage({ source: source.document, story, mode, resource, offset });
        const edited = driver.session.captureSaveSnapshot().document;
        for (const serializationMode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
          const saved = await serializeCanonicalSave({
            snapshot: driver.session.captureSaveSnapshot(),
            options: { mode: serializationMode },
            featureFlags: { selectiveSave: true },
          });
          const zip = await JSZip.loadAsync(saved.buffer);
          const sourceBuffer = source.document.originalBuffer ?? panic("Missing parsed source");
          expect(await canonicalSaveParagraphXml(saved.buffer, "76000002")).toBe(
            await canonicalSaveParagraphXml(sourceBuffer, "76000002"),
          );
          const sourceZip = await JSZip.loadAsync(sourceBuffer);
          const ownedPart = await owningPart(sourceZip, story);
          const writable = new Set([
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "[Content_Types].xml",
            "docProps/core.xml",
            ownedPart,
            relationshipPath(ownedPart),
          ]);
          if (
            canonicalJson(edited.package.styles) !== canonicalJson(source.document.package.styles)
          )
            writable.add("word/styles.xml");
          await Promise.all(
            Object.entries(sourceZip.files).map(async ([entry, file]) => {
              if (file.dir || writable.has(entry)) return;
              expect(await zip.file(entry)?.async("uint8array")).toEqual(
                await file.async("uint8array"),
              );
            }),
          );
          await assertSavedOwner(zip, edited, OP_STORIES.MAIN);
          if (story !== OP_STORIES.MAIN) await assertSavedOwner(zip, edited, story);
          const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
          for (const [alias, media] of reopened.package.media ?? []) {
            const primary =
              reopened.package.media?.get(media.path) ?? panic("Media alias has no primary path");
            expect(new Uint8Array(media.data)).toEqual(new Uint8Array(primary.data));
            expect(media.mimeType).toBe(primary.mimeType);
            if (alias !== media.path) expect(alias).toBe(media.path.replace(/^word\//u, ""));
          }
          for (const ownedStory of story === OP_STORIES.MAIN ? [story] : [OP_STORIES.MAIN, story]) {
            const expected =
              drawings(edited, ownedStory).at(0)?.image ?? panic("Expected copied image missing");
            const actual =
              drawings(reopened, ownedStory).at(0)?.image ?? panic("Reopened copied image missing");
            expect(actual.size).toEqual(expected.size);
            expect(actual.wrap.type).toBe(expected.wrap.type);
            if (resource === "http") expect(actual?.src).toBeUndefined();
            else expect(actual?.src).toBe(expected?.src);
          }
          const expectedPackage = exportedPackageOf(edited);
          expect(describePackageDifferences(expectedPackage, reopened)).toEqual({
            messages: [],
            omitted: 0,
          });
          const repeated = await serializeCanonicalSave({
            snapshot: driver.session.captureSaveSnapshot(),
            options: { mode: serializationMode },
            featureFlags: { selectiveSave: true },
          });
          await assertSavedOwner(await JSZip.loadAsync(repeated.buffer), edited, story);
          assertExactModel(driver.session.captureSaveSnapshot().document, edited);
        }
        if (story !== OP_STORIES.MAIN) {
          const undo = driver.session.prepareUndo(driver.storyState, story).unwrap();
          driver.storyState = driver.storyState.apply(undo.transaction);
          undo.publish().unwrap();
        }
        const undo = driver.session.prepareUndo(driver.bodyState).unwrap();
        driver.bodyState = driver.bodyState.apply(undo.transaction);
        undo.publish().unwrap();
        assertExactModel(driver.session.captureSaveSnapshot().document, driver.before);
        const redo = driver.session.prepareRedo(driver.bodyState).unwrap();
        driver.bodyState = driver.bodyState.apply(redo.transaction);
        redo.publish().unwrap();
        if (story !== OP_STORIES.MAIN) {
          const redoStory = driver.session.prepareRedo(driver.storyState, story).unwrap();
          driver.storyState = driver.storyState.apply(redoStory.transaction);
          redoStory.publish().unwrap();
        }
        assertExactModel(driver.session.captureSaveSnapshot().document, edited);
      }),
      {
        seed: 20261022,
        numRuns: 2,
        id: "copied $resource image saves with exact $kind ownership in $mode",
      },
    );
  },
);

test.each(Object.values(FOLIO_DOCX_SERIALIZATION_MODE))(
  "copied image conflicting owner refuses %s save",
  async (mode) => {
    const story = STORIES.header;
    const source = await destination({ story, conflictingBinding: "present" });
    const driver = await copyImage({
      source: source.document,
      story,
      mode: "editing",
      resource: "dataUrl",
      offset: 0,
    });
    expect(driver.allocated.image.rId).toBe(source.conflictId);
    const snapshot = driver.session.captureSaveSnapshot();
    const before = snapshot.document;
    const part = await owningPart(
      await JSZip.loadAsync(source.document.originalBuffer ?? panic("Missing parsed baseline")),
      story,
    );
    const result = await Result.tryPromise({
      try: () =>
        serializeCanonicalSave({
          snapshot,
          options: { mode },
          featureFlags: { selectiveSave: true },
        }),
      catch: (error: unknown) => error,
    });
    expect(result.isErr()).toBe(true);
    if (result.isOk()) panic("Conflicting owning relationship unexpectedly saved");
    expect(result.error).toBeInstanceOf(CanonicalImageRelationshipRefusalError);
    expect(result.error).toMatchObject({
      gap: CANONICAL_GAP.save,
      reason: "conflict",
      part: relationshipPath(part),
      relationshipId: source.conflictId,
    });
    assertExactModel(driver.session.captureSaveSnapshot().document, before);
  },
);

test("parser-derived enrichment preserves every authored field and the equality oracle detects its mutation", async () => {
  const source = await destination({ story: OP_STORIES.MAIN });
  const driver = await copyImage({
    source: source.document,
    story: OP_STORIES.MAIN,
    mode: "editing",
    resource: "dataUrl",
    offset: 0,
  });
  const authored = {
    id: "12345",
    filename: "authored.png",
    mimeType: "image/png",
    pictureNames: { name: "Authored picture", alt: "Alternative", title: "Authored title" },
  } satisfies Required<Pick<Image, keyof ReturnType<typeof imageDerivedFactsOf>>>;
  const document = structuredClone(driver.session.captureSaveSnapshot().document);
  const image =
    drawings(document, OP_STORIES.MAIN).at(0)?.image ?? panic("Missing authored picture");
  Object.assign(image, authored);
  const expected = exportedPackageOf(document);
  expect(drawings(expected, OP_STORIES.MAIN).at(0)?.image).toMatchObject(authored);
  for (const [field, value] of Object.entries(authored)) {
    const changed = structuredClone(expected);
    const changedImage =
      drawings(changed, OP_STORIES.MAIN).at(0)?.image ?? panic("Missing changed picture");
    Reflect.set(
      changedImage,
      field,
      typeof value === "string" ? `${value}-changed` : { ...value, name: "Changed" },
    );
    expect(describePackageDifferences(expected, changed).messages.length).toBeGreaterThan(0);
  }
});
