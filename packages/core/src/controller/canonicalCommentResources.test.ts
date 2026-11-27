/**
 * A copied comment's images and external links paste with it: their ids
 * name the comments part's own relationships, so the paste imports each by
 * its bytes or target and the save writes it into the comments part's
 * relationships, where a reopen resolves it again.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { validateDocxPackage } from "@stll/docx-core";
import { ZipArchive } from "@stll/docx-core/zip";

import { assertExactModel } from "../../../../test/exactModel";
import { propertyTestTimeout } from "../../../../test/property-testing";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { parseRelationships, RELATIONSHIP_TYPES } from "../docx/relsParser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { Comment, Document, Hyperlink, Image } from "../types/document";
import { readModel } from "../types/documentModel";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { prepareCanonicalPaste } from "./canonicalClipboard";

setDefaultTimeout(propertyTestTimeout(60_000));

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NAMESPACES = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const RELATIONSHIPS_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";

// 1x1 transparent PNG.
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
const COMMENT_LINK = "https://comment.example.test/";

const seedDocument = (text: string): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [{ type: "run", content: [{ type: "text", text }] }],
        },
      ],
    },
  },
});

const picture = `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const relationships = (entries: string): string =>
  `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NS}">${entries}</Relationships>`;

/**
 * A package whose comment 3 holds a picture and comment 4 an external link,
 * both by ids of the comments part's relationships. The document part names
 * other targets under the same ids.
 */
const sourcePackage = async (): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(seedDocument("Source")));
  const contentTypes = await zip.entry("[Content_Types].xml")?.text();
  if (contentTypes === undefined)
    throw new TypeError("Expected the seed package's content types.");
  zip.set(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
    ),
  );
  // The document part's own rId1 and rId2 name other targets.
  const documentRels = await zip.entry("word/_rels/document.xml.rels")?.text();
  if (documentRels === undefined || !documentRels.includes('Id="rId1"'))
    throw new TypeError("Expected the seed document's relationships.");
  const otherLink = documentRels.includes('Id="rId2"')
    ? ""
    : `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="https://document.example.test/" TargetMode="External"/>`;
  zip.set(
    "word/_rels/document.xml.rels",
    documentRels.replace(
      "</Relationships>",
      `${otherLink}<Relationship Id="rId900" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/></Relationships>`,
    ),
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NAMESPACES}><w:body><w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Body</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r><w:r><w:commentReference w:id="4"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments ${NAMESPACES}><w:comment w:id="3" w:author="Source"><w:p><w:r><w:t xml:space="preserve">Pictured </w:t></w:r>${picture}</w:p></w:comment><w:comment w:id="4" w:author="Source"><w:p><w:hyperlink r:id="rId2"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p></w:comment></w:comments>`,
  );
  zip.set(
    "word/_rels/comments.xml.rels",
    relationships(
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/comment.png"/><Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="${COMMENT_LINK}" TargetMode="External"/>`,
    ),
  );
  zip.set(
    "word/media/comment.png",
    Uint8Array.from(atob(PNG_BASE64), (char) => char.charCodeAt(0)),
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const imagesOf = (comment: Comment): Image[] =>
  comment.content.flatMap((paragraph) =>
    paragraph.content.flatMap((item) =>
      item.type === "run"
        ? item.content.flatMap((child) =>
            child.type === "drawing" ? [child.image] : [],
          )
        : [],
    ),
  );

const linksOf = (comment: Comment): Hyperlink[] =>
  comment.content.flatMap((paragraph) =>
    paragraph.content.flatMap((item) =>
      item.type === "hyperlink" ? [item] : [],
    ),
  );

/** The one comment holding text `text`, or a thrown TypeError. */
const commentHolding = (
  comments: readonly Comment[] | undefined,
  text: string,
): Comment => {
  const found = (comments ?? []).filter((comment) =>
    JSON.stringify(comment.content).includes(text),
  );
  const only = found.at(0);
  if (only === undefined || found.length !== 1)
    throw new TypeError(`Expected one comment holding "${text}".`);
  return only;
};

const pictureSource = (comment: Comment): string => {
  const image = imagesOf(comment).at(0);
  if (image?.src === undefined)
    throw new TypeError("Expected a comment picture with its bytes.");
  return image.src;
};

const linkTarget = (comment: Comment): string => {
  const link = linksOf(comment).at(0);
  if (link?.href === undefined)
    throw new TypeError("Expected a comment link with its target.");
  return link.href;
};

const destinationDocument = async (): Promise<Document> =>
  parseDocx(await createDocx(seedDocument("Destination")), {
    preloadFonts: false,
  });

/** Every relationship id the comments part names, with its relationship. */
const commentRelationships = async (zip: ZipArchive) => {
  const commentsXml = await zip.entry("word/comments.xml")?.text();
  const relsXml = await zip.entry("word/_rels/comments.xml.rels")?.text();
  if (commentsXml === undefined || relsXml === undefined)
    throw new TypeError("Expected the comments part and its relationships.");
  const rels = parseRelationships(relsXml);
  return [...commentsXml.matchAll(/r:(?:embed|id)="(?<id>[^"]+)"/gu)].map(
    (match) => {
      const id = match.groups?.["id"];
      if (id === undefined) throw new TypeError("Expected a relationship id.");
      return { id, relationship: rels.get(id) };
    },
  );
};

describe("copied comments holding pictures and links", () => {
  test("the comments part's own relationships resolve its pictures and links", async () => {
    const source = await parseDocx(await sourcePackage(), {
      preloadFonts: false,
    });
    const comments = source.package.document.comments;
    expect(pictureSource(commentHolding(comments, "Pictured"))).toStartWith(
      "data:image/png",
    );
    expect(linkTarget(commentHolding(comments, "linked"))).toBe(COMMENT_LINK);
  });

  test("without the source package's parts, a copied comment's kept markup naming an external target carries its relationship into another package", async () => {
    const sourceDocument = await parseDocx(await sourcePackage(), {
      preloadFonts: false,
    });
    const linked = commentHolding(
      sourceDocument.package.document.comments,
      "linked",
    );
    linked.preserved = {
      children: [
        {
          index: 1,
          xml: `<w:p xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:hyperlink r:id="rId2"/></w:p>`,
        },
      ],
    };
    const session = createCanonicalSession(
      await destinationDocument(),
    ).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    const before = session.document;
    const prepared = prepareCanonicalPaste({
      session,
      state,
      slice: new Slice(
        Fragment.from(
          schema.node("paragraph", null, [
            schema.node("commentReference", { commentId: 4 }),
          ]),
        ),
        1,
        1,
      ),
      sourceDocument,
    });
    if (prepared.isErr()) throw prepared.error;
    expect(prepared.value.warnings).toEqual([]);
    const next = publishCanonicalProjection({
      session,
      state,
      commit: prepared.value,
    }).unwrap().state;
    // The kept link names a fresh relationship of the comments part, to the
    // target its source named.
    const pasted = readModel(session.document);
    const kept = pasted.package.document.comments
      ?.flatMap((comment) => comment.preserved?.children ?? [])
      .at(0);
    if (kept === undefined) throw new TypeError("The kept markup is missing.");
    const id = /r:id="(?<id>[^"]+)"/u.exec(kept.xml)?.groups?.["id"];
    if (id === undefined) throw new TypeError("The kept markup names no id.");
    expect(
      pasted.package.annotationRelationships?.get("comments")?.get(id)?.target,
    ).toBe(COMMENT_LINK);
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      options: { mode: FOLIO_DOCX_SERIALIZATION_MODE.full },
    });
    const zip = await ZipArchive.read(saved.buffer);
    const commentRels = parseRelationships(
      (await zip.entry("word/_rels/comments.xml.rels")?.text()) ?? "",
    );
    expect(commentRels.get(id)?.target).toBe(COMMENT_LINK);
    publishCanonicalProjection({
      session,
      state: next,
      commit: session.prepareUndo(next).unwrap(),
    }).unwrap();
    assertExactModel(session.document, before);
  });

  for (const suggesting of [false, true])
    test(`a paste ${suggesting ? "suggesting" : "editing"} carries them, the save writes them into the comments part, and undo removes them`, async () => {
      const sourceDocument = await parseDocx(await sourcePackage(), {
        preloadFonts: false,
      });
      const sourceComments = sourceDocument.package.document.comments;
      const expectedPicture = pictureSource(
        commentHolding(sourceComments, "Pictured"),
      );
      const destination = await destinationDocument();
      const session = createCanonicalSession(destination).unwrap();
      session.setMode(
        suggesting
          ? { type: "suggesting", author: "Clipboard" }
          : { type: "editing" },
      );
      let state = EditorState.create({ schema, doc: session.projection.doc });
      const slice = new Slice(
        Fragment.from(
          schema.node("paragraph", null, [
            schema.text("Copied"),
            schema.node("commentReference", { commentId: 3 }),
            schema.node("commentReference", { commentId: 4 }),
          ]),
        ),
        1,
        1,
      );
      const before = session.document;
      const prepared = prepareCanonicalPaste({
        session,
        state,
        slice,
        sourceDocument,
      });
      if (prepared.isErr()) throw prepared.error;
      state = publishCanonicalProjection({
        session,
        state,
        commit: prepared.value,
      }).unwrap().state;

      const pasted = readModel(session.document).package.document.comments;
      expect(pictureSource(commentHolding(pasted, "Pictured"))).toBe(
        expectedPicture,
      );
      expect(linkTarget(commentHolding(pasted, "linked"))).toBe(COMMENT_LINK);

      const snapshot = session.captureSaveSnapshot();
      for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
        const saved = await serializeCanonicalSave({
          snapshot,
          featureFlags: { selectiveSave: true },
          options: { mode },
        });
        expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual(
          { valid: true },
        );
        const zip = await ZipArchive.read(saved.buffer);
        const named = await commentRelationships(zip);
        expect(named.length).toBeGreaterThanOrEqual(2);
        for (const { relationship } of named) {
          if (relationship === undefined)
            throw new TypeError(
              "A comment names a relationship its part lacks.",
            );
          if (relationship.type === RELATIONSHIP_TYPES.image)
            expect(
              zip.entry(`word/${relationship.target}`),
              `Saved media for ${relationship.target} (${mode})`,
            ).toBeDefined();
          else {
            expect(relationship.type).toBe(RELATIONSHIP_TYPES.hyperlink);
            expect(relationship.target).toBe(COMMENT_LINK);
          }
        }
        const reopened = await parseDocx(saved.buffer, {
          preloadFonts: false,
        });
        const reopenedComments = reopened.package.document.comments;
        expect(
          pictureSource(commentHolding(reopenedComments, "Pictured")),
        ).toBe(expectedPicture);
        expect(linkTarget(commentHolding(reopenedComments, "linked"))).toBe(
          COMMENT_LINK,
        );
      }

      expect(session.canUndo).toBe(true);
      publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap();
      assertExactModel(session.document, before);
    });
});
