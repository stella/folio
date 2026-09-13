import {expect, test} from "bun:test";
import JSZip from "jszip";

import {FolioDocxReviewer} from "../ai-edits/headless";
import {createDocx} from "../docx/rezip";
import {parseRelationships, RELATIONSHIP_TYPES, resolveRelativePath} from "../docx/relsParser";
import {createEmptyDocument} from "../utils/createDocument";
import {compareDocx} from "./compare";

const PNG = Uint8Array.fromBase64("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

test("Folio-exact imports an embedded watermark without replacing base media", async () => {
  const baseZip = await JSZip.loadAsync(await createDocx(createEmptyDocument()));
  const originalMedia = new Uint8Array([1, 2, 3]);
  baseZip.file("word/media/image1.png", originalMedia);
  baseZip.file("word/media/folio-import-1.png", originalMedia);
  const base = await baseZip.generateAsync({type: "arraybuffer"});
  const document = createEmptyDocument();
  document.package.document.finalSectionProperties = {...document.package.document.finalSectionProperties, headerReferences: [{type: "default", rId: "rId_header"}]};
  document.package.headers = new Map([["rId_header", {type: "header", hdrFtrType: "default", content: [{type: "paragraph", content: []}]}]]);
  const zip = await JSZip.loadAsync(await createDocx(document));
  zip.file("word/header1.xml", `<w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape id="WordPictureWatermark1" style="position:absolute;width:72pt;height:72pt"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p></w:hdr>`);
  zip.file("word/_rels/header1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/></Relationships>`);
  zip.file("word/media/image1.png", PNG);
  const contentTypes = zip.file("[Content_Types].xml");
  if (!contentTypes) throw new Error("Missing content types");
  zip.file("[Content_Types].xml", (await contentTypes.async("text")).replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'));
  const result = await compareDocx(base, await zip.generateAsync({type: "arraybuffer"}), {author: "Reviewer", timestamp: "2026-09-13T00:00:00.000Z", revisionFormat: "folio-exact"});
  if (result.isErr()) throw result.error;
  expect(result.value.verification.status).toBe("verified");
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
    const bytes = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(bytes);
    const headers = reopened.toDocument().package.headers;
    expect(headers?.size ?? 0).toBe(decision === "accept" ? 1 : 0);
    const saved = await JSZip.loadAsync(bytes);
    if (decision === "reject") {
      expect(Object.keys(saved.files).filter((name) => name.startsWith("word/media/folio-import-") && name !== "word/media/folio-import-1.png")).toEqual([]);
      continue;
    }
    expect(await saved.file("word/media/image1.png")?.async("uint8array")).toEqual(originalMedia);
    expect(await saved.file("word/media/folio-import-1.png")?.async("uint8array")).toEqual(originalMedia);
    const rels = saved.file("word/_rels/header1.xml.rels");
    if (!rels) throw new Error("Missing imported header relationships");
    const image = [...parseRelationships(await rels.async("text")).values()].find(({type}) => type === RELATIONSHIP_TYPES.image);
    if (!image) throw new Error("Missing watermark image relationship");
    const media = saved.file(resolveRelativePath("word/_rels/header1.xml.rels", image.target));
    if (!media) throw new Error("Missing imported image");
    expect(await media.async("uint8array")).toEqual(PNG);
    expect(headers?.values().next().value?.rawWatermarkXml).toContain('width:72pt;height:72pt');
  }
});
