/**
 * Every attribute a range marker arrived with must survive a real save.
 *
 * The corpus found the instance: `w:moveFromRangeStart` and
 * `w:moveToRangeStart` reached the disk without the `w:author` their schema
 * type requires, so Word repaired the document. The class is wider — each of
 * these markers declared its own subset of `CT_MarkupRange` and so each lost a
 * different part of it — and only a property over the whole attribute set can
 * hold it closed.
 *
 * The test lives at the scripts level because the oracle does: the schema
 * validator behind the corpus gate is `scripts/lib/corpus-schema-validator.ts`,
 * and a package may not import from `scripts`. It checks two things a
 * round-trip comparison alone would not: that every generated attribute comes
 * back, and that the part folio wrote gains no schema violation on the marker.
 *
 * The saved paragraph is edited before the save so the verbatim capture cannot
 * replay it. Replay is the path that hid this defect from every other test.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Paragraph } from "@stll/folio-core/types/document";

import { propertyConfig, propertyTestTimeout } from "../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
import {
  loadSchemaGraph,
  type SchemaViolation,
  validateOoxmlPart,
} from "./lib/corpus-schema-validator";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * The attributes each marker's schema type declares, required and optional.
 *
 * `CT_MarkupRange` for an end marker and a comment range, `CT_Bookmark` for a
 * bookmark start, `CT_MoveBookmark` for a move-range start.
 */
const MARKUP_RANGE = ["id", "displacedByCustomXml"] as const;
const BOOKMARK_RANGE = [...MARKUP_RANGE, "colFirst", "colLast", "name"] as const;
const MOVE_BOOKMARK = [...BOOKMARK_RANGE, "author", "date"] as const;

const MARKER_ATTRIBUTES = {
  bookmarkStart: BOOKMARK_RANGE,
  bookmarkEnd: MARKUP_RANGE,
  moveFromRangeStart: MOVE_BOOKMARK,
  moveFromRangeEnd: MARKUP_RANGE,
  moveToRangeStart: MOVE_BOOKMARK,
  moveToRangeEnd: MARKUP_RANGE,
  commentRangeStart: MARKUP_RANGE,
  commentRangeEnd: MARKUP_RANGE,
} as const;

type MarkerName = keyof typeof MARKER_ATTRIBUTES;

const MARKER_NAMES = Object.keys(MARKER_ATTRIBUTES) as MarkerName[];

/** A comment id the fixture defines, so a comment range is not an orphan. */
const COMMENT_ID = 7;

/** ST_String values that survive without normalisation: no trimming, no escaping surprises. */
const plainText = (maxLength: number) =>
  fc
    .stringMatching(/^[A-Za-z0-9_.-]+$/u)
    .filter((value) => value.length > 0 && value.length <= maxLength);

const isoDate = fc
  .integer({ min: 0, max: 2_000_000_000 })
  .map((seconds) => `${new Date(seconds * 1000).toISOString().slice(0, 19)}Z`);

const attributeValue = (name: string): fc.Arbitrary<string> => {
  switch (name) {
    case "id":
      // Below `MAX_REVISION_ID`: a larger value is folded on save by design.
      return fc.integer({ min: 0, max: 0x7fff_fffe }).map(String);
    case "displacedByCustomXml":
      return fc.constantFrom("next", "prev");
    case "colFirst":
    case "colLast":
      return fc.integer({ min: 0, max: 63 }).map(String);
    case "date":
      return isoDate;
    default:
      return plainText(24);
  }
};

type MarkerCase = { marker: MarkerName; attributes: Record<string, string> };

/**
 * A comment range needs its comment, or the reconciliation drops the marker as
 * an orphan and the property would pass by testing nothing. Its `w:id` is the
 * comment's, so it is the one attribute the case cannot choose freely.
 */
const isCommentRange = (marker: MarkerName): boolean => marker.startsWith("commentRange");

const markerCase = (marker: MarkerName): fc.Arbitrary<MarkerCase> =>
  fc
    .record(
      Object.fromEntries(
        MARKER_ATTRIBUTES[marker].map((name) => [
          name,
          name === "id" && isCommentRange(marker)
            ? fc.constant(String(COMMENT_ID))
            : attributeValue(name),
        ]),
      ) as Record<string, fc.Arbitrary<string>>,
    )
    .map((attributes) => ({ marker, attributes }));

const markerXml = ({ marker, attributes }: MarkerCase): string => {
  const written = Object.entries(attributes)
    .map(([name, value]) => `w:${name}="${value}"`)
    .join(" ");
  return `<w:${marker} ${written}/>`;
};

const TRACKED = 'w:id="900" w:author="A" w:date="2024-01-01T00:00:00Z"';

/**
 * The subject marker inside a construct folio will not normalise away.
 *
 * An unpaired move-range marker is removed as unbalanced, and an orphan
 * comment range is dropped, so each case writes the whole range and puts the
 * generated attributes on the one marker under test.
 */
const markerGroup = (subject: MarkerCase): string => {
  const written = markerXml(subject);
  const id = subject.attributes["id"] ?? "0";
  const inside = "<w:r><w:t>inside</w:t></w:r>";
  switch (subject.marker) {
    case "bookmarkStart":
      return `${written}${inside}<w:bookmarkEnd w:id="${id}"/>`;
    case "bookmarkEnd":
      return `<w:bookmarkStart w:id="${id}" w:name="bm"/>${inside}${written}`;
    case "moveFromRangeStart":
      return `${written}<w:moveFrom ${TRACKED}><w:r><w:delText>gone</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="${id}"/>`;
    case "moveFromRangeEnd":
      return `<w:moveFromRangeStart w:id="${id}" w:name="mv" w:author="A" w:date="2024-01-01T00:00:00Z"/><w:moveFrom ${TRACKED}><w:r><w:delText>gone</w:delText></w:r></w:moveFrom>${written}`;
    case "moveToRangeStart":
      return `${written}<w:moveTo ${TRACKED}><w:r><w:t>here</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="${id}"/>`;
    case "moveToRangeEnd":
      return `<w:moveToRangeStart w:id="${id}" w:name="mv" w:author="A" w:date="2024-01-01T00:00:00Z"/><w:moveTo ${TRACKED}><w:r><w:t>here</w:t></w:r></w:moveTo>${written}`;
    case "commentRangeStart":
      return `${written}${inside}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
    case "commentRangeEnd":
      return `<w:commentRangeStart w:id="${id}"/>${inside}${written}<w:r><w:commentReference w:id="${id}"/></w:r>`;
    default: {
      const unreachable: never = subject.marker;
      return unreachable;
    }
  }
};

const buildDocx = async (subject: MarkerCase): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const commented = isCommentRange(subject.marker);
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>` +
      `<w:p><w:r><w:t>before</w:t></w:r>${markerGroup(subject)}<w:r><w:t>after</w:t></w:r></w:p>` +
      `<w:sectPr/></w:body></w:document>`,
  );
  if (commented) {
    zip.file(
      "word/comments.xml",
      `${XML_DECLARATION}<w:comments xmlns:w="${W_NAMESPACE}">` +
        `<w:comment w:id="${COMMENT_ID}" w:author="A" w:date="2024-01-01T00:00:00Z">` +
        `<w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`,
    );
    const types = await zip.file("[Content_Types].xml")?.async("text");
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    if (types === undefined || rels === undefined) {
      throw new Error("the empty package lost its packaging parts");
    }
    zip.file(
      "[Content_Types].xml",
      types.replace(
        "</Types>",
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
      ),
    );
    zip.file(
      "word/_rels/document.xml.rels",
      rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdPropertyComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
      ),
    );
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

/** The attributes of the first `w:<marker>` element in the saved part. */
const savedAttributes = (xml: string, marker: MarkerName): Record<string, string> | undefined => {
  const element = new RegExp(`<w:${marker}\\b([^>]*)/>`, "u").exec(xml);
  if (element === null) {
    return undefined;
  }
  const attributes: Record<string, string> = {};
  for (const [, name, value] of element[1]?.matchAll(/\bw:([\w]+)="([^"]*)"/gu) ?? []) {
    if (name !== undefined && value !== undefined) {
      attributes[name] = value;
    }
  }
  return attributes;
};

/**
 * Edit the paragraph so nothing about it can be replayed from captured bytes.
 *
 * folio replays a paragraph's captured markup whenever the model still agrees
 * with it, and replay preserves every attribute for free. The defect only
 * exists on the serializer path, so the property has to take it.
 */
const editFirstParagraph = (paragraph: Paragraph): void => {
  paragraph.content.push({ type: "run", content: [{ type: "text", text: "edited" }] });
};

type SaveResult = { documentXml: string; violations: SchemaViolation[] };

type SaveOptions = { subject: MarkerCase; viaEditor: boolean };

const saveEdited = async ({ subject, viaEditor }: SaveOptions): Promise<SaveResult> => {
  const parsed = await parseDocx(await buildDocx(subject), { preloadFonts: false });
  const document = viaEditor ? fromProseDoc(toProseDoc(parsed), parsed) : parsed;
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph") {
    throw new Error("the fixture did not parse as a paragraph");
  }
  editFirstParagraph(first);
  const saved = await repackDocx(document, { updateModifiedDate: false });
  const zip = await JSZip.loadAsync(saved);
  const documentXml = (await zip.file("word/document.xml")?.async("text")) ?? "";
  return {
    documentXml,
    violations: validateOoxmlPart({ graph: await loadSchemaGraph(), xml: documentXml }),
  };
};

/**
 * How much of a marker the editor projection carries.
 *
 * An attribute can be lost on either leg, so modelling one the editor drops
 * only trades a parse-time loss for a round-trip loss. `bookmarkBoundary` is a
 * real ProseMirror node, so a bookmark survives the editor whole. A comment
 * range is rebuilt from the comment mark, which carries the id and nothing
 * else. A move range has no projection at all and its markers do not survive
 * unless a content control replays them verbatim. The last two are losses the
 * public corpus already records; stating them here makes closing one a test
 * change rather than a silent one.
 */
const EDITOR_PROJECTIONS = { whole: "whole", idOnly: "id-only", dropped: "dropped" } as const;

type EditorProjection = (typeof EDITOR_PROJECTIONS)[keyof typeof EDITOR_PROJECTIONS];

const EDITOR_PROJECTION = {
  bookmarkStart: EDITOR_PROJECTIONS.whole,
  bookmarkEnd: EDITOR_PROJECTIONS.whole,
  commentRangeStart: EDITOR_PROJECTIONS.idOnly,
  commentRangeEnd: EDITOR_PROJECTIONS.idOnly,
  moveFromRangeStart: EDITOR_PROJECTIONS.dropped,
  moveFromRangeEnd: EDITOR_PROJECTIONS.dropped,
  moveToRangeStart: EDITOR_PROJECTIONS.dropped,
  moveToRangeEnd: EDITOR_PROJECTIONS.dropped,
} as const satisfies Record<MarkerName, EditorProjection>;

describe("range markers keep every attribute a real save re-serializes", () => {
  for (const marker of MARKER_NAMES) {
    test(`w:${marker}`, async () => {
      await fc.assert(
        fc.asyncProperty(markerCase(marker), async (subject) => {
          const { documentXml, violations } = await saveEdited({ subject, viaEditor: false });

          const written = savedAttributes(documentXml, marker);
          expect(written).toBeDefined();
          expect(written).toEqual(subject.attributes);

          // The marker must gain no violation of its own, and it must not
          // make the part invalid anywhere else either.
          expect(violations).toEqual([]);
        }),
        // Each run writes and re-reads a package, so the budget is per-marker.
        propertyConfig({ numRuns: 20 }),
      );
    });
  }
});

describe("what the editor projection carries, it carries whole", () => {
  for (const marker of MARKER_NAMES) {
    test(`w:${marker}`, async () => {
      const projection = EDITOR_PROJECTION[marker];
      await fc.assert(
        fc.asyncProperty(markerCase(marker), async (subject) => {
          const { documentXml, violations } = await saveEdited({ subject, viaEditor: true });
          const written = savedAttributes(documentXml, marker);
          expect(violations).toEqual([]);
          switch (projection) {
            case EDITOR_PROJECTIONS.whole:
              expect(written).toEqual(subject.attributes);
              return;
            case EDITOR_PROJECTIONS.idOnly:
              expect(written).toEqual({ id: subject.attributes["id"] ?? "" });
              return;
            case EDITOR_PROJECTIONS.dropped:
              expect(written).toBeUndefined();
              return;
            default: {
              const unreachable: never = projection;
              return unreachable;
            }
          }
        }),
        propertyConfig({ numRuns: 20 }),
      );
    });
  }
});
