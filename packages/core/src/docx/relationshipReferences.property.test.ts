/**
 * A forced serialization never invents a relationship target.
 *
 * `r:id`, `r:embed` and the header reference are the same construct three
 * times: an NCName that names a relationship, which may be valid, absent,
 * present but empty, or dangling. Three of those four name nothing, and the
 * defect class is treating them as a lookup key anyway — an empty `r:embed`
 * that falls back to `rId1` binds a picture to whatever the first relationship
 * happens to be, which in a Word package is `styles.xml`.
 *
 * The property generates packages over all four cases and runs them through a
 * save with every rebuildable capture removed, so the real serializers run for
 * every block rather than replaying the bytes the parse captured. What must
 * hold is that each reference resolves, after the round trip, to exactly what
 * it resolved to before: a target to the same target, and nothing to nothing.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { BlockContent, Document, RelationshipMap } from "../types/document";
import { parseDocx } from "./parser";
import { resolveRelationshipId } from "./relsParser";
import { repackDocx } from "./rezip";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Relationship ids the generated package defines. */
const STYLES_RID = "rId1";
const IMAGE_RID = "rId2";
const HYPERLINK_RID = "rId3";
const HEADER_RID = "rId4";
/** An id the package deliberately never defines. */
const MISSING_RID = "rId90";

/** The four states a relationship reference can be in. */
const REFERENCE_STATES = ["valid", "absent", "empty", "dangling"] as const;
type ReferenceState = (typeof REFERENCE_STATES)[number];

/** `r:<name>="..."`, or nothing at all when the reference is absent. */
const referenceAttribute = (name: string, state: ReferenceState, validId: string): string => {
  switch (state) {
    case "valid":
      return ` r:${name}="${validId}"`;
    case "absent":
      return "";
    case "empty":
      return ` r:${name}=""`;
    case "dangling":
      return ` r:${name}="${MISSING_RID}"`;
    default:
      return state satisfies never;
  }
};

type BodyItem =
  | { kind: "text"; text: string }
  | { kind: "picture"; state: ReferenceState }
  /** A `wp:inline` with no `a:graphic`: a chart or an OLE frame minimises to this. */
  | { kind: "graphicless" }
  | { kind: "hyperlink"; state: ReferenceState; text: string };

const bodyItemXml = (item: BodyItem): string => {
  switch (item.kind) {
    case "text":
      return `<w:p><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r></w:p>`;
    case "picture": {
      const embed = referenceAttribute("embed", item.state, IMAGE_RID);
      return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip${embed}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    }
    case "graphicless":
      return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Chart 2"/></wp:inline></w:drawing></w:r></w:p>`;
    case "hyperlink": {
      const id = referenceAttribute("id", item.state, HYPERLINK_RID);
      return `<w:p><w:hyperlink${id}><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r></w:hyperlink></w:p>`;
    }
    default:
      return item satisfies never;
  }
};

const DOCUMENT_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  `xmlns:r="${OFFICE_RELATIONSHIPS}"`,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

type PackageSpec = { items: readonly BodyItem[]; headerReference: ReferenceState };

const documentXml = ({ items, headerReference }: PackageSpec): string => {
  const header = `<w:headerReference w:type="default"${referenceAttribute("id", headerReference, HEADER_RID)}/>`;
  return `${XML_DECL}<w:document ${DOCUMENT_NAMESPACES}><w:body>${items
    .map(bodyItemXml)
    .join(
      "",
    )}<w:sectPr>${header}<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
};

/** A one-pixel PNG, so the image relationship has real bytes behind it. */
const PNG_BYTES = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  ),
  (character) => character.codePointAt(0) ?? 0,
);

const buildPackage = async (spec: PackageSpec): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file("word/document.xml", documentXml(spec));
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="${STYLES_RID}" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/><Relationship Id="${IMAGE_RID}" Type="${OFFICE_RELATIONSHIPS}/image" Target="media/image1.png"/><Relationship Id="${HYPERLINK_RID}" Type="${OFFICE_RELATIONSHIPS}/hyperlink" Target="https://example.invalid/a" TargetMode="External"/><Relationship Id="${HEADER_RID}" Type="${OFFICE_RELATIONSHIPS}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECL}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECL}<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file("word/media/image1.png", PNG_BYTES);
  return await zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Every capture slot the model can rebuild, cleared.
 *
 * Mirrors the corpus gate's `reserialize` invariant, and goes one slot further:
 * a drawing's `rawXml` is cleared too, because the claim under test is about
 * what the serializer writes when it has only the model to write it from.
 */
const CAPTURE_SLOTS = new Set([
  "sourceXml",
  "gridSourceXml",
  "verbatimXml",
  "verbatimFingerprint",
  "rawPropertiesXml",
  "rawEndPropertiesXml",
  "rawImageFingerprint",
  "rawXml",
  "rawXmlMode",
  "rawWatermarkXml",
]);

const clearCaptures = (value: unknown, seen: WeakSet<object>): void => {
  if (Array.isArray(value)) {
    for (const item of value) clearCaptures(item, seen);
    return;
  }
  if (typeof value !== "object" || value === null || seen.has(value)) return;
  seen.add(value);
  const record = value as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (CAPTURE_SLOTS.has(key)) {
      record[key] = undefined;
      continue;
    }
    clearCaptures(record[key], seen);
  }
};

const withoutCaptures = (document: Document): Document => {
  const cloned = structuredClone(document);
  clearCaptures(cloned.package, new WeakSet());
  return cloned;
};

/** What one reference in the model resolves to, as a comparable string. */
const resolution = (
  relationships: RelationshipMap | undefined,
  rId: string | undefined,
): string => {
  const resolved = resolveRelationshipId(relationships, rId);
  switch (resolved.status) {
    case "absent":
      return "absent";
    case "dangling":
      return "dangling";
    case "resolved":
      return `${resolved.relationship.type}|${resolved.relationship.target}`;
    default:
      return resolved satisfies never;
  }
};

/**
 * Every relationship reference the body holds, in document order.
 *
 * Targets rather than ids: a save may renumber a relationship, and renumbering
 * is not the defect. Naming a different part is.
 */
const referenceProjection = (document: Document): string[] => {
  const { relationships } = document.package;
  const projected: string[] = [];
  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") continue;
      for (const item of block.content) {
        if (item.type === "hyperlink") {
          projected.push(`hyperlink:${resolution(relationships, item.rId)}`);
          continue;
        }
        if (item.type !== "run") continue;
        for (const child of item.content) {
          if (child.type !== "drawing") continue;
          projected.push(
            `drawing:${resolution(relationships, child.image.rId)}:${child.image.filename ?? "-"}`,
          );
        }
      }
    }
  };
  visitBlocks(document.package.document.content);
  for (const reference of document.package.document.finalSectionProperties.headerReferences ?? []) {
    projected.push(`header:${reference.type}:${resolution(relationships, reference.rId)}`);
  }
  return projected;
};

/**
 * No output reference names a relationship the source did not define.
 *
 * `referenceProjection` compares targets the model resolved; this compares the
 * raw attributes the serializer wrote, so an id folio invented is caught even
 * if the reparse happens to read it back as nothing.
 */
const writtenReferences = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = (await zip.file("word/document.xml")?.async("text")) ?? "";
  return [...xml.matchAll(/\sr:(?:id|embed|link)="([^"]*)"/gu)].map(([, id]) => id ?? "");
};

const referenceStateArbitrary = fc.constantFrom(...REFERENCE_STATES);

const bodyItemArbitrary: fc.Arbitrary<BodyItem> = fc.oneof(
  fc.constant<BodyItem>({ kind: "graphicless" }),
  referenceStateArbitrary.map<BodyItem>((state) => ({ kind: "picture", state })),
  referenceStateArbitrary.map<BodyItem>((state) => ({
    kind: "hyperlink",
    state,
    text: "linked text",
  })),
  fc.constant<BodyItem>({ kind: "text", text: "plain paragraph" }),
);

const packageArbitrary: fc.Arbitrary<PackageSpec> = fc.record({
  items: fc.array(bodyItemArbitrary, { minLength: 1, maxLength: 6 }),
  headerReference: referenceStateArbitrary,
});

describe("relationship references under a forced serialization (property)", () => {
  test(
    "resolve after the round trip to exactly what they resolved to before",
    async () => {
      await fc.assert(
        fc.asyncProperty(packageArbitrary, async (spec) => {
          const source = await parseDocx(await buildPackage(spec), { preloadFonts: false });
          const saved = await repackDocx(withoutCaptures(source), { updateModifiedDate: false });
          const reparsed = await parseDocx(saved, { preloadFonts: false });

          expect(referenceProjection(reparsed)).toEqual(referenceProjection(source));

          const defined = new Set(source.package.relationships?.keys() ?? []);
          for (const written of await writtenReferences(saved)) {
            expect(defined.has(written) || written === MISSING_RID).toBe(true);
          }
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
