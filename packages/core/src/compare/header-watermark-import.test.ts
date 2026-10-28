import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  parseRelationships,
  RELATIONSHIP_TYPES,
  resolveRelativePath,
} from "../docx/relsParser";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const PNG = Uint8Array.fromBase64(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
);
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

test("Folio-exact imports an embedded watermark without replacing base media", async () => {
  const baseZip = await ZipArchive.read(
    await createDocx(createEmptyDocument()),
  );
  const originalMedia = new Uint8Array([1, 2, 3]);
  baseZip.set("word/media/image1.png", originalMedia);
  baseZip.set("word/media/folio-import-1.png", originalMedia);
  const base = (await baseZip.generate({ compression: "store" })).buffer;
  const document = createEmptyDocument();
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: "rId_header" }],
  };
  document.package.headers = new Map([
    [
      "rId_header",
      {
        type: "header",
        hdrFtrType: "default",
        content: [{ type: "paragraph", content: [] }],
      },
    ],
  ]);
  const zip = await ZipArchive.read(await createDocx(document));
  zip.set(
    "word/header1.xml",
    `<w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape id="WordPictureWatermark1" style="position:absolute;width:72pt;height:72pt"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p></w:hdr>`,
  );
  zip.set(
    "word/_rels/header1.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/></Relationships>`,
  );
  zip.set("word/media/image1.png", PNG);
  const contentTypes = zip.entry("[Content_Types].xml");
  if (!contentTypes) throw new Error("Missing content types");
  zip.set(
    "[Content_Types].xml",
    (await contentTypes.text()).replace(
      "</Types>",
      '<Default Extension="png" ContentType="image/png"/></Types>',
    ),
  );
  const result = await compareDocx(
    base,
    (await zip.generate({ compression: "store" })).buffer,
    {
      author: "Reviewer",
      timestamp: "2026-09-13T00:00:00.000Z",
      revisionFormat: "folio-exact",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification.status).toBe("verified");
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    const bytes = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(bytes);
    const headers = reopened.toDocument().package.headers;
    expect(headers?.size ?? 0).toBe(decision === "accept" ? 1 : 0);
    const saved = await ZipArchive.read(bytes);
    if (decision === "reject") {
      expect(
        saved
          .paths()
          .filter(
            (name) =>
              name.startsWith("word/media/folio-import-") &&
              name !== "word/media/folio-import-1.png",
          ),
      ).toEqual([]);
      continue;
    }
    expect(await saved.entry("word/media/image1.png")?.bytes()).toEqual(
      originalMedia,
    );
    expect(await saved.entry("word/media/folio-import-1.png")?.bytes()).toEqual(
      originalMedia,
    );
    const rels = saved.entry("word/_rels/header1.xml.rels");
    if (!rels) throw new Error("Missing imported header relationships");
    const image = [...parseRelationships(await rels.text()).values()].find(
      ({ type }) => type === RELATIONSHIP_TYPES.image,
    );
    if (!image) throw new Error("Missing watermark image relationship");
    const media = saved.entry(
      resolveRelativePath("word/_rels/header1.xml.rels", image.target),
    );
    if (!media) throw new Error("Missing imported image");
    expect(await media.bytes()).toEqual(PNG);
    expect(headers?.values().next().value?.rawWatermarkXml).toContain(
      "width:72pt;height:72pt",
    );
  }
});
