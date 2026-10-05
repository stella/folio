import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { parseDocx } from "./parser";
import { parseRelationships, RELATIONSHIP_TYPES } from "./relsParser";
import { createEmptyDocx, DocxPackageFidelityError, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));
const RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const HEADER_RELS = "word/_rels/header1.xml.rels";
const HEADER = "word/header1.xml";
const IMAGE = "word/media/watermark.png";
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
type RelationshipFixtureOptions = {
  requestedId: string;
  state: "free" | "matching" | "conflicting" | "unresolved";
  prefixed: boolean;
};
const fixture = async ({ requestedId, state, prefixed }: RelationshipFixtureOptions) => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const types = await zip.file("[Content_Types].xml")?.async("text");
  zip.file(
    "[Content_Types].xml",
    types?.replace(
      "</Types>",
      '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ) ?? "",
  );
  const documentRels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  zip.file(
    "word/_rels/document.xml.rels",
    documentRels?.replace(
      "</Relationships>",
      `<Relationship Id="rIdHeader99" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/></Relationships>`,
    ) ?? "",
  );
  zip.file(
    "word/document.xml",
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Untouched body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader99"/></w:sectPr></w:body></w:document>',
  );
  zip.file(
    HEADER,
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p w14:paraId="00000002"/></w:hdr>',
  );
  const prefix = prefixed ? "p:" : "";
  const root = prefixed ? `xmlns:p="${RELS_NAMESPACE}"` : `xmlns="${RELS_NAMESPACE}"`;
  const foreign = `<${prefix}Relationship Id="rIdOpaque" Type="urn:foreign" Target="../custom/opaque.bin" opaque="authored"/>`;
  const occupied =
    state === "matching" || state === "conflicting"
      ? `<${prefix}Relationship Id="${requestedId}" Type="${RELATIONSHIP_TYPES.image}" Target="media/${state === "matching" ? "watermark" : "other"}.png"/>`
      : "";
  const alias = `<${prefix}Relationship Id="rIdAlias" Type="${RELATIONSHIP_TYPES.image}" Target="media/watermark.png"/>`;
  const rels = `<${prefix}Relationships ${root}>${foreign}${occupied}${alias}</${prefix}Relationships>`;
  zip.file(HEADER_RELS, rels);
  zip.file(IMAGE, PNG, { base64: true });
  zip.file("word/media/other.png", PNG, { base64: true });
  zip.file("custom/opaque.bin", new Uint8Array([17, 29, 41]));
  const source = await zip.generateAsync({ type: "arraybuffer" });
  const document = await parseDocx(source, { preloadFonts: false });
  const header = document.package.headers?.get("rIdHeader99");
  if (!header) throw new Error("Missing header fixture");
  header.watermark = {
    kind: "picture",
    imageRId: requestedId,
    imageTarget: state === "unresolved" ? "word/media/missing.png" : IMAGE,
  };
  return { document, sourceZip: zip, rels, foreign };
};

test("generated canonical watermark relationships preserve free and matching ids and refuse conflicting or unresolved targets", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.integer({ min: 10, max: 100_000 }),
      fc.boolean(),
      async (suffix, prefixed) => {
        const requestedId = `rId${suffix}`;
        for (const state of ["free", "matching", "conflicting", "unresolved"] as const) {
          const input = await fixture({ requestedId, state, prefixed });
          if (state === "conflicting" || state === "unresolved") {
            await expect(
              repackDocx(input.document, { bodyAuthority: "canonical", updateModifiedDate: false }),
            ).rejects.toBeInstanceOf(DocxPackageFidelityError);
            expect(input.document.package.headers?.get("rIdHeader99")?.watermark).toMatchObject({
              imageRId: requestedId,
            });
            continue;
          }
          const saved = await repackDocx(input.document, {
            bodyAuthority: "canonical",
            updateModifiedDate: false,
          });
          const zip = await JSZip.loadAsync(saved);
          const rels = await zip.file(HEADER_RELS)?.async("text");
          expect(rels).toContain(input.foreign);
          if (state === "matching") expect(rels).toBe(input.rels);
          const requested = parseRelationships(rels ?? "").get(requestedId);
          expect(requested?.target).toBe("media/watermark.png");
          expect(parseRelationships(rels ?? "").get("rIdAlias")?.target).toBe(
            "media/watermark.png",
          );
          const reopened = await parseDocx(saved, { preloadFonts: false });
          expect(reopened.package.headers?.get("rIdHeader99")?.watermark).toMatchObject({
            kind: "picture",
            imageRId: requestedId,
            imageTarget: IMAGE,
          });
          expect(reopened.package.document.content).toStrictEqual(
            input.document.package.document.content,
          );
          for (const path of [
            IMAGE,
            "word/media/other.png",
            "custom/opaque.bin",
            "word/styles.xml",
          ]) {
            expect(await zip.file(path)?.async("uint8array")).toEqual(
              await input.sourceZip.file(path)?.async("uint8array"),
            );
          }
        }
      },
    ),
    { seed: 20261019, numRuns: 12 },
  );
});

test("default watermark rebinding still reuses an existing matching relationship", async () => {
  const input = await fixture({ requestedId: "rIdRequested", state: "free", prefixed: false });
  const saved = await repackDocx(input.document, { updateModifiedDate: false });
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(reopened.package.headers?.get("rIdHeader99")?.watermark).toMatchObject({
    kind: "picture",
    imageRId: "rIdAlias",
    imageTarget: IMAGE,
  });
  const zip = await JSZip.loadAsync(saved);
  expect(await zip.file(HEADER_RELS)?.async("text")).toBe(input.rels);
});
