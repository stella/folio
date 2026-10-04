/** Authored package-registry spellings, derived from the existing operation seed. */
import JSZip from "jszip";
import { createDocx } from "@stll/folio-core/docx/rezip";
import type { Document } from "../../../packages/docx-core/src/model/document";
import type { OpSeed } from "../../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { panic } from "better-result";

export const operationPackageBytes = async (
  document: Document,
  seed: OpSeed,
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createDocx(structuredClone(document)));
  const file = zip.file("[Content_Types].xml");
  if (!file) return panic("A generated package needs content types.");
  const xml = await file.async("text");
  const entries = Array.from(
    xml.matchAll(/<(?:Default|Override)\b[^>]*\/>/gu),
    (match) => match[0],
  );
  if (seed.first % 2 === 0) entries.reverse();
  const prefix = seed.second % 2 === 0 ? "ct:" : "";
  const extension = seed.third % 2 === 0 ? "opaque" : "OPAQUE";
  const extra = seed.inherit
    ? `<Default Extension="${extension}" ContentType="application/octet-stream"/>`
    : '<Override PartName="/custom/keep.opaque" ContentType="application/octet-stream"/>';
  entries.splice(seed.depth % (entries.length + 1), 0, extra);
  const gap = seed.first % 3 === 0 ? "\r\n  " : "\n\t";
  const body = entries.map((entry) => entry.replace(/^</u, `<${prefix}`)).join(gap);
  let source = `<?xml version="1.0" encoding="UTF-8"?>${gap}<${prefix}Types xmlns${prefix ? ":ct" : ""}="http://schemas.openxmlformats.org/package/2006/content-types">${gap}${body}${gap}</${prefix}Types>`;
  if (seed.zeroWidth !== undefined) source = source.replaceAll('"', "'");
  zip.file("[Content_Types].xml", source);
  const authoredParagraphParts = zip.file(/^word\/(?:document|header[^/]*|footer[^/]*)\.xml$/u);
  const headerParts = authoredParagraphParts.filter((part) => /^word\/header/u.test(part.name));
  const footerParts = authoredParagraphParts.filter((part) => /^word\/footer/u.test(part.name));
  if (headerParts.length !== (document.package.headers?.size ?? 0))
    return panic("Every generated header needs a package XML part.");
  if (footerParts.length !== (document.package.footers?.size ?? 0))
    return panic("Every generated footer needs a package XML part.");
  for (const part of authoredParagraphParts) {
    const paragraphXml = await part.async("text");
    zip.file(part.name, paragraphXml.replaceAll("<w:pPr>", `<w:pPr>${gap}`));
  }
  zip.file("custom/keep.opaque", "unrelated package bytes");
  return zip.generateAsync({ type: "arraybuffer" });
};
