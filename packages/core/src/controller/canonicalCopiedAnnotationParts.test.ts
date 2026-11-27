/**
 * Kept markup in a copied footnote or comment (a chart with its embedded
 * workbook, kept block markup) names the parts it reaches by the note or
 * comment part's own relationships. Pasted into the package it came from it
 * resolves where it is; pasted into another, the parts are copied under fresh
 * names, the ids are rewritten to fresh relationships of the destination's
 * note or comment part, and the save writes them there: every relationship
 * resolves after a reopen, and an undo leaves none of the copies in the save.
 */
import { describe, expect, test } from "bun:test";
import type { Node as ProseNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import {
  parseRelationships,
  RELATIONSHIP_TYPES,
  resolveRelativePath,
} from "../docx/relsParser";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  parseXmlDocument,
} from "../docx/xmlParser";
import type { Document } from "../types/document";
import { readModel } from "../types/documentModel";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { decodeText } from "../utils/textEncoding";
import { prepareCanonicalPaste } from "./canonicalClipboard";
import {
  readClipboardPackageParts,
  sliceAnnotationMarkup,
  sliceKeptMarkup,
} from "./canonicalClipboardParts";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NAMESPACES =
  `xmlns:w="${W_NS}" xmlns:r="${R_NS}" ` +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const RELATIONSHIPS_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const CHART_TYPE = `${R_NS}/chart`;
const CHUNK_TYPE = `${R_NS}/aFChunk`;
const CHART_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const chart = (id: string, docPr: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>` +
  `<wp:docPr id="${String(docPr)}" name="Chart ${String(docPr)}"/><a:graphic>` +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  `<c:chart r:id="${id}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const chunk = (id: string) => `<w:altChunk r:id="${id}"/>`;

const relationships = (
  entries: readonly { id: string; type: string; target: string }[],
) =>
  `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NS}">${entries
    .map(
      ({ id, type, target }) =>
        `<Relationship Id="${id}" Type="${type}" Target="${target}"${target.startsWith("https:") ? ' TargetMode="External"' : ""}/>`,
    )
    .join("")}</Relationships>`;

const separators =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

/** The chart parts every package here holds, by path. */
const CHART_PARTS: Record<string, { type: string; body: string }> = {
  "word/charts/chart1.xml": {
    type: CHART_CONTENT_TYPE,
    body: `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}"><c:chart/><c:externalData r:id="rId1"/></c:chartSpace>`,
  },
  "word/charts/_rels/chart1.xml.rels": {
    type: "application/vnd.openxmlformats-package.relationships+xml",
    body: relationships([
      {
        id: "rId1",
        type: `${R_NS}/package`,
        target: "../embeddings/workbook1.xlsx",
      },
    ]),
  },
  "word/embeddings/workbook1.xlsx": {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    body: "PK workbook bytes",
  },
  "word/chunk1.html": {
    type: "text/html",
    body: "<html><body><p>Chunk</p></body></html>",
  },
};

type PackageOptions = {
  /** The leading digit of the package's paragraph ids. */
  paraIdBase: number;
  /** The name the footnotes part takes, `footnotes` or a differently cased one. */
  footnotesName: string;
  body: string;
  footnotes: string;
  footnoteRelationships: readonly {
    id: string;
    type: string;
    target: string;
  }[];
  comments?: string;
  commentRelationships?: readonly {
    id: string;
    type: string;
    target: string;
  }[];
  parts: Record<string, { type: string; body: string }>;
};

/** A package with a footnotes part, optionally a comments part, and `parts`. */
const packageOf = async ({
  paraIdBase,
  footnotesName,
  body,
  footnotes,
  footnoteRelationships,
  comments,
  commentRelationships = [],
  parts,
}: PackageOptions): Promise<Document> => {
  const zip = new ZipArchive();
  // Each paragraph takes an id of its own across the package's parts.
  let paragraphs = 0;
  const identified = (xml: string) =>
    xml.replaceAll("<w:p>", () => {
      paragraphs += 1;
      return `<w:p w14:paraId="${String(paraIdBase * 10_000_000 + paragraphs)}">`;
    });
  const overrides = Object.entries(parts)
    .filter(([path]) => !path.endsWith(".rels"))
    .map(
      ([path, { type }]) =>
        `<Override PartName="/${path}" ContentType="${type}"/>`,
    )
    .join("");
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      `<Override PartName="/word/${footnotesName}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>` +
      (comments === undefined
        ? ""
        : '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>') +
      `${overrides}</Types>`,
  );
  zip.set(
    "_rels/.rels",
    relationships([
      {
        id: "rId1",
        type: RELATIONSHIP_TYPES.officeDocument,
        target: "word/document.xml",
      },
    ]),
  );
  zip.set(
    "word/_rels/document.xml.rels",
    relationships([
      {
        id: "rId1",
        type: RELATIONSHIP_TYPES.footnotes,
        target: `${footnotesName}.xml`,
      },
      ...(comments === undefined
        ? []
        : [
            {
              id: "rId2",
              type: RELATIONSHIP_TYPES.comments,
              target: "comments.xml",
            },
          ]),
    ]),
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NAMESPACES}><w:body>${identified(body)}<w:sectPr/></w:body></w:document>`,
  );
  zip.set(
    `word/${footnotesName}.xml`,
    `${XML_DECLARATION}<w:footnotes ${NAMESPACES}>${identified(separators + footnotes)}</w:footnotes>`,
  );
  zip.set(
    `word/_rels/${footnotesName}.xml.rels`,
    relationships(footnoteRelationships),
  );
  if (comments !== undefined) {
    zip.set(
      "word/comments.xml",
      `${XML_DECLARATION}<w:comments ${NAMESPACES}>${identified(comments)}</w:comments>`,
    );
    zip.set(
      "word/_rels/comments.xml.rels",
      relationships(commentRelationships),
    );
  }
  for (const [path, part] of Object.entries(parts)) zip.set(path, part.body);
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

/**
 * The source: a paragraph referencing footnote 1 and comment 5, each holding
 * a chart and kept block markup by its own part's ids, which the main
 * document's relationships hold for other parts.
 */
const source = () =>
  packageOf({
    paraIdBase: 1,
    footnotesName: "footnotes",
    body:
      '<w:p><w:r><w:t xml:space="preserve">Copied </w:t></w:r>' +
      '<w:r><w:footnoteReference w:id="1"/></w:r>' +
      '<w:commentRangeStart w:id="5"/><w:r><w:t>noted</w:t></w:r>' +
      '<w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p>' +
      "<w:p><w:r><w:t>After</w:t></w:r></w:p>",
    footnotes:
      `<w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">Note </w:t></w:r>${chart("rId1", 31)}</w:p>` +
      `${chunk("rId2")}<w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:footnote>`,
    footnoteRelationships: [
      { id: "rId1", type: CHART_TYPE, target: "charts/chart1.xml" },
      { id: "rId2", type: CHUNK_TYPE, target: "chunk1.html" },
    ],
    comments:
      `<w:comment w:id="5" w:author="Reviewer"><w:p><w:r><w:t xml:space="preserve">Comment </w:t></w:r>${chart("rId2", 41)}</w:p>` +
      `${chunk("rId1")}</w:comment>`,
    commentRelationships: [
      { id: "rId1", type: CHUNK_TYPE, target: "chunk1.html" },
      { id: "rId2", type: CHART_TYPE, target: "charts/chart1.xml" },
    ],
    parts: CHART_PARTS,
  });

/**
 * Another package: its footnotes part is cased differently and already holds
 * a note whose chart takes the ids the copies must not, and a chart part of
 * its own under the source chart's name.
 */
const destination = () =>
  packageOf({
    paraIdBase: 2,
    footnotesName: "Footnotes",
    body:
      '<w:p><w:r><w:t xml:space="preserve">Target </w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
      "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>",
    footnotes: `<w:footnote w:id="1"><w:p><w:r><w:t>Own</w:t></w:r>${chart("rId1", 7)}</w:p></w:footnote>`,
    footnoteRelationships: [
      { id: "rId1", type: CHART_TYPE, target: "charts/chart1.xml" },
      {
        id: "rId2",
        type: RELATIONSHIP_TYPES.hyperlink,
        target: "https://own.example.test/",
      },
    ],
    parts: {
      "word/charts/chart1.xml": {
        type: CHART_CONTENT_TYPE,
        body: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
      },
    },
  });

/** The copied range: the source's first paragraph, with its note and comment. */
const copiedSlice = (doc: ProseNode) => doc.slice(0, doc.child(0).nodeSize);

type Pasted = {
  session: CanonicalSession;
  state: EditorState;
  before: Document;
};

const paste = async (
  into: "same" | "other",
  mode: "editing" | "suggesting",
): Promise<Pasted> => {
  // Within one package the copy is the session's own document, as the
  // editor reads it.
  const copy = createCanonicalSession(await source()).unwrap();
  const copiedFrom = readModel(copy.document);
  const slice = copiedSlice(copy.projection.doc);
  const sourceParts = await readClipboardPackageParts({
    document: copiedFrom,
    markup: sliceKeptMarkup(slice),
    annotationMarkup: sliceAnnotationMarkup(slice, copiedFrom),
  });
  const session =
    into === "same"
      ? copy
      : createCanonicalSession(await destination()).unwrap();
  session.setMode(
    mode === "editing"
      ? { type: "editing" }
      : { type: "suggesting", author: "Paster" },
  );
  let state = EditorState.create({ doc: session.projection.doc });
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, state.doc.child(0).nodeSize - 1),
    ),
  );
  const before = readModel(session.document);
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: copiedFrom,
    ...(sourceParts === undefined ? {} : { sourceParts }),
  });
  if (prepared.isErr()) throw prepared.error;
  const next = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  return { session, state: next, before };
};

const save = async (session: CanonicalSession, mode: "full" | "selective") =>
  (
    await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      options: {
        mode:
          mode === "full"
            ? FOLIO_DOCX_SERIALIZATION_MODE.full
            : FOLIO_DOCX_SERIALIZATION_MODE.preferSelective,
      },
      featureFlags: { selectiveSave: true },
    })
  ).buffer;

const textOf = async (zip: ZipArchive, path: string) =>
  decodeText((await zip.entry(path)?.bytes()) ?? new Uint8Array(), {
    ignoreBOM: true,
  });

/** The content type the package registry gives `path`. */
const contentTypeOf = (registry: string, path: string) => {
  const elements = getChildElements(parseXmlDocument(registry));
  for (const element of elements)
    if (
      getLocalName(element.name ?? "") === "Override" &&
      getAttribute(element, null, "PartName")?.toLowerCase() ===
        `/${path}`.toLowerCase()
    )
      return getAttribute(element, null, "ContentType");
  const extension = path.slice(path.lastIndexOf(".") + 1);
  for (const element of elements)
    if (
      getLocalName(element.name ?? "") === "Default" &&
      getAttribute(element, null, "Extension") === extension
    )
      return getAttribute(element, null, "ContentType");
  return null;
};

/** Each relationship id an office attribute names in `xml`. */
const namedIds = (xml: string) =>
  [...xml.matchAll(/\sr:id="(?<id>[^"]+)"/gu)].flatMap(({ groups }) =>
    groups?.["id"] === undefined ? [] : [groups["id"]],
  );

/**
 * Assert every internal relationship of every part resolves to a part with a
 * content type, and every id the part at `partPath` names resolves in its own
 * relationships: the targets, by type, each id reaches.
 */
const resolvedIn = async (zip: ZipArchive, partPath: string) => {
  const registry = await textOf(zip, "[Content_Types].xml");
  const paths = new Set(zip.paths());
  const relsParts = await Promise.all(
    zip
      .paths()
      .filter((path) => path.endsWith(".rels"))
      .map(async (path) => ({ path, xml: await textOf(zip, path) })),
  );
  for (const { path: relsPath, xml } of relsParts)
    for (const relationship of parseRelationships(xml).values()) {
      if (relationship.targetMode === "External") continue;
      const target = resolveRelativePath(relsPath, relationship.target);
      expect(paths.has(target), `${relsPath} ${relationship.id}`).toBe(true);
      expect(contentTypeOf(registry, target), target).not.toBeNull();
    }
  const slash = partPath.lastIndexOf("/");
  const relsPath = `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
  const held = parseRelationships(await textOf(zip, relsPath));
  const targets: { id: string; type: string; target: string }[] = [];
  for (const id of namedIds(await textOf(zip, partPath))) {
    const relationship = held.get(id);
    if (relationship === undefined)
      throw new TypeError(`${partPath} names ${id}, which its part lacks.`);
    targets.push({
      id,
      type: relationship.type.slice(relationship.type.lastIndexOf("/") + 1),
      target:
        relationship.targetMode === "External"
          ? relationship.target
          : resolveRelativePath(relsPath, relationship.target),
    });
  }
  return targets;
};

/** The targets, by type, the ids `partPath` names reach. */
const targetsByType = (
  targets: readonly { type: string; target: string }[],
) => {
  const byType = new Map<string, string[]>();
  for (const { type, target } of targets)
    byType.set(type, [...(byType.get(type) ?? []), target].toSorted());
  return byType;
};

/** The names the copies take beside the destination's own parts. */
const COPIED_PATHS: ReadonlySet<string> = new Set([
  "word/charts/chart2.xml",
  "word/charts/_rels/chart2.xml.rels",
  "word/embeddings/workbook1.xlsx",
  "word/chunk1.html",
]);

describe.each(["editing", "suggesting"] as const)("in %s", (mode) => {
  test.each(["full", "selective"] as const)(
    "a footnote and a comment pasted from another package bring their parts into their own parts' relationships (%s save)",
    async (saveMode) => {
      const { session, state, before } = await paste("other", mode);
      const pasted = readModel(session.document);
      expect(pasted.package.footnotes?.length).toBe(2);
      expect(pasted.package.document.comments?.length).toBe(1);
      // The pasted note's and comment's ids name fresh relationships of the
      // destination's own parts, not the ids the destination's note holds.
      const footnoteRelationships =
        pasted.package.annotationRelationships?.get("footnotes");
      expect(footnoteRelationships?.get("rId1")?.target).toBe(
        "charts/chart1.xml",
      );
      expect(footnoteRelationships?.size).toBe(4);

      const zip = await ZipArchive.read(await save(session, saveMode));
      // The footnotes part keeps its casing, and so do its relationships.
      expect(zip.paths()).toContain("word/Footnotes.xml");
      expect(zip.paths()).toContain("word/_rels/Footnotes.xml.rels");
      expect(zip.paths()).not.toContain("word/footnotes.xml");
      expect(zip.paths()).not.toContain("word/_rels/footnotes.xml.rels");

      const footnoteTargets = await resolvedIn(zip, "word/Footnotes.xml");
      const byType = targetsByType(footnoteTargets);
      // The destination's own chart and the copied one, each its own part.
      expect(byType.get("chart")).toEqual([
        "word/charts/chart1.xml",
        "word/charts/chart2.xml",
      ]);
      expect(byType.get("aFChunk")).toEqual(["word/chunk1.html"]);
      expect(await textOf(zip, "word/charts/chart1.xml")).not.toContain(
        "externalData",
      );
      expect(await textOf(zip, "word/charts/chart2.xml")).toBe(
        CHART_PARTS["word/charts/chart1.xml"]?.body ?? "",
      );
      expect(
        [
          ...parseRelationships(
            await textOf(zip, "word/charts/_rels/chart2.xml.rels"),
          ).values(),
        ].map(({ target }) =>
          resolveRelativePath("word/charts/_rels/chart2.xml.rels", target),
        ),
      ).toEqual(["word/embeddings/workbook1.xlsx"]);

      // The comment's chart and chunk reach the same copies: one source part
      // is copied once, whichever part's markup names it.
      const commentTargets = targetsByType(
        await resolvedIn(zip, "word/comments.xml"),
      );
      expect(commentTargets.get("chart")).toEqual(["word/charts/chart2.xml"]);
      expect(commentTargets.get("aFChunk")).toEqual(["word/chunk1.html"]);
      const registry = await textOf(zip, "[Content_Types].xml");
      expect(contentTypeOf(registry, "word/charts/chart2.xml")).toBe(
        CHART_CONTENT_TYPE,
      );
      expect(contentTypeOf(registry, "word/chunk1.html")).toBe("text/html");

      // Reopened, the note and the comment still name their parts.
      const reopened = await parseDocx(
        (await zip.generate({ compression: "store" })).buffer,
        { preloadFonts: false },
      );
      const reopenedNotes = JSON.stringify(reopened.package.footnotes);
      const reopenedComments = JSON.stringify(
        reopened.package.document.comments,
      );
      for (const { id } of footnoteTargets) {
        expect(reopenedNotes).toContain(`r:id=\\"${id}\\"`);
        expect(
          reopened.package.annotationRelationships?.get("footnotes")?.has(id),
        ).toBe(true);
      }
      expect(reopenedComments).toContain("c:chart");
      expect(reopenedComments).toContain("w:altChunk");

      // Undone, the save holds none of the copies.
      publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap();
      const undone = readModel(session.document);
      expect(undone.package.copiedParts ?? new Map()).toEqual(new Map());
      expect(undone.package.annotationRelationships).toEqual(
        before.package.annotationRelationships,
      );
      const restored = await ZipArchive.read(await save(session, saveMode));
      await resolvedIn(restored, "word/Footnotes.xml");
      expect(restored.paths().filter((path) => COPIED_PATHS.has(path))).toEqual(
        [],
      );
      expect(
        parseRelationships(
          await textOf(restored, "word/_rels/Footnotes.xml.rels"),
        ).size,
      ).toBe(2);
    },
  );

  test.each(["full", "selective"] as const)(
    "a footnote and a comment pasted within their package resolve where they are (%s save)",
    async (saveMode) => {
      const { session, state, before } = await paste("same", mode);
      const pasted = readModel(session.document);
      expect(pasted.package.footnotes?.length).toBe(2);
      expect(pasted.package.document.comments?.length).toBe(2);
      // Nothing is copied: the ids name what they named.
      expect(pasted.package.copiedParts ?? new Map()).toEqual(new Map());
      expect(pasted.package.annotationRelationships).toEqual(
        before.package.annotationRelationships,
      );
      const zip = await ZipArchive.read(await save(session, saveMode));
      const footnoteTargets = targetsByType(
        await resolvedIn(zip, "word/footnotes.xml"),
      );
      expect(footnoteTargets.get("chart")).toEqual([
        "word/charts/chart1.xml",
        "word/charts/chart1.xml",
      ]);
      expect(footnoteTargets.get("aFChunk")).toEqual([
        "word/chunk1.html",
        "word/chunk1.html",
      ]);
      const commentTargets = targetsByType(
        await resolvedIn(zip, "word/comments.xml"),
      );
      expect(commentTargets.get("chart")).toEqual([
        "word/charts/chart1.xml",
        "word/charts/chart1.xml",
      ]);
      expect(zip.paths()).not.toContain("word/charts/chart2.xml");

      publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap();
      expect(readModel(session.document).package.footnotes?.length).toBe(1);
    },
  );
});

test("without the source package's parts, a copied note's kept markup naming them refuses in another package", async () => {
  const copiedFrom = readModel(
    createCanonicalSession(await source()).unwrap().document,
  );
  const slice = copiedSlice(
    createCanonicalSession(copiedFrom).unwrap().projection.doc,
  );
  const session = createCanonicalSession(await destination()).unwrap();
  const state = EditorState.create({ doc: session.projection.doc });
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: copiedFrom,
  });
  if (prepared.isOk()) throw new TypeError("The paste was expected to refuse.");
  expect(prepared.error.ground).toEqual({ gap: "story-content-projection" });
});

test("a copy reads the parts its notes' and comments' markup reaches through their own parts' relationships", async () => {
  const copiedFrom = readModel(
    createCanonicalSession(await source()).unwrap().document,
  );
  const slice = copiedSlice(
    createCanonicalSession(copiedFrom).unwrap().projection.doc,
  );
  const annotationMarkup = sliceAnnotationMarkup(slice, copiedFrom);
  expect([...annotationMarkup.keys()].toSorted()).toEqual([
    "comments",
    "footnotes",
  ]);
  const parts = await readClipboardPackageParts({
    document: copiedFrom,
    markup: sliceKeptMarkup(slice),
    annotationMarkup,
  });
  expect([...(parts?.parts.keys() ?? [])].toSorted()).toEqual(
    Object.keys(CHART_PARTS).toSorted(),
  );
});
