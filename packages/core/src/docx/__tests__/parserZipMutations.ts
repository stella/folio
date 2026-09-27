import path from "node:path";

import { escapeXmlAttribute } from "@stll/docx-core";
import JSZip from "jszip";

export const ZIP_MUTATIONS = [
  "drop-part",
  "rename-part",
  "duplicate-part",
  "truncate",
  "bad-crc",
  "wrong-content-type",
  "missing-relationship-target",
] as const;

export type ZipMutation = (typeof ZIP_MUTATIONS)[number];

const decodeUtf8 = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const requirePart = (zip: JSZip, part: string): JSZip.JSZipObject => {
  const file = zip.file(part);
  if (file === null) {
    throw new Error(`ZIP mutation target does not exist: ${part}`);
  }
  return file;
};

const generated = (zip: JSZip): Promise<Uint8Array> =>
  zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

const packagePathForRelationship = (
  relationshipFile: string,
  target: string,
): string | undefined => {
  if (target.startsWith("/") || /^[a-z][a-z\d+.-]*:/iu.test(target)) {
    return undefined;
  }
  const ownerDirectory =
    relationshipFile === "_rels/.rels"
      ? ""
      : path.posix.dirname(path.posix.dirname(relationshipFile));
  return path.posix.normalize(path.posix.join(ownerDirectory, target));
};

const wrongContentType = async (source: Uint8Array, targetPart: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(source);
  requirePart(zip, targetPart);
  const contentTypes = await requirePart(zip, "[Content_Types].xml").async("string");
  const partName = `/${targetPart.replace(/^\/+/, "")}`;
  const overridePattern = /<Override\b[^>]*>/gu;
  let matched = false;
  const updated = contentTypes.replace(overridePattern, (tag) => {
    const name = tag.match(/\bPartName\s*=\s*(["'])(.*?)\1/u)?.[2];
    if (name !== partName) {
      return tag;
    }
    matched = true;
    return /\bContentType\s*=\s*(["']).*?\1/u.test(tag)
      ? tag.replace(/\bContentType\s*=\s*(["']).*?\1/u, 'ContentType="application/x-folio-fuzz"')
      : tag.replace(/\s*\/?>$/u, ' ContentType="application/x-folio-fuzz"/>');
  });
  if (!matched) {
    const closing = "</Types>";
    if (!updated.includes(closing)) {
      throw new Error("ZIP mutation could not find the [Content_Types].xml closing tag");
    }
    zip.file(
      "[Content_Types].xml",
      updated.replace(
        closing,
        `<Override PartName="${escapeXmlAttribute(partName)}" ContentType="application/x-folio-fuzz"/>${closing}`,
      ),
    );
  } else {
    zip.file("[Content_Types].xml", updated);
  }
  return generated(zip);
};

const missingRelationshipTarget = async (
  source: Uint8Array,
  targetPart: string,
): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(source);
  const wanted = path.posix.normalize(targetPart.replace(/^\/+/, ""));
  for (const [relationshipFile, entry] of Object.entries(zip.files)) {
    if (entry.dir || !relationshipFile.endsWith(".rels")) {
      continue;
    }
    const xml = await entry.async("string");
    let changed = false;
    const updated = xml.replace(/<Relationship\b[^>]*>/gu, (tag) => {
      const target = tag.match(/\bTarget\s*=\s*(["'])(.*?)\1/u)?.[2];
      if (target === undefined || packagePathForRelationship(relationshipFile, target) !== wanted) {
        return tag;
      }
      changed = true;
      const missingTarget = `${target}.missing`;
      return tag.replace(
        /\bTarget\s*=\s*(["']).*?\1/u,
        `Target="${escapeXmlAttribute(missingTarget)}"`,
      );
    });
    if (changed) {
      zip.file(relationshipFile, updated);
      return generated(zip);
    }
  }
  throw new Error(`ZIP mutation found no relationship targeting part: ${targetPart}`);
};

const badCentralDirectoryCrc = (source: Uint8Array, targetPart: string): Uint8Array => {
  const bytes = source.slice();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  while (offset + 46 <= bytes.length) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      offset += 1;
      continue;
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const end = offset + 46 + nameLength + extraLength + commentLength;
    if (end > bytes.length) {
      break;
    }
    const name = decodeUtf8(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (name === targetPart) {
      view.setUint32(offset + 16, view.getUint32(offset + 16, true) ^ 1, true);
      return bytes;
    }
    offset = end;
  }
  throw new Error(`ZIP mutation could not find central directory entry for: ${targetPart}`);
};

const duplicateCentralDirectoryEntry = (source: Uint8Array, targetPart: string): Uint8Array => {
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
  const minimumEocd = 22;
  const eocdSearchStart = Math.max(0, source.byteLength - minimumEocd - 0xffff);
  let eocdOffset = -1;
  for (let offset = source.byteLength - minimumEocd; offset >= eocdSearchStart; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      const commentLength = view.getUint16(offset + 20, true);
      if (offset + minimumEocd + commentLength === source.byteLength) {
        eocdOffset = offset;
        break;
      }
    }
  }
  if (eocdOffset < 0) {
    throw new Error("ZIP mutation could not find the end-of-central-directory record");
  }

  const entriesOnDisk = view.getUint16(eocdOffset + 8, true);
  const entryCount = view.getUint16(eocdOffset + 10, true);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  if (
    entriesOnDisk === 0xffff ||
    entryCount === 0xffff ||
    entriesOnDisk >= 0xfffe ||
    entryCount >= 0xfffe ||
    centralDirectorySize === 0xffffffff ||
    centralDirectoryOffset === 0xffffffff ||
    view.getUint16(eocdOffset + 4, true) !== 0 ||
    view.getUint16(eocdOffset + 6, true) !== 0 ||
    centralDirectoryEnd !== eocdOffset
  ) {
    throw new Error(
      "ZIP mutation requires a non-ZIP64 archive with an unextended central directory",
    );
  }

  let offset = centralDirectoryOffset;
  while (offset + 46 <= centralDirectoryEnd) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("ZIP mutation found an invalid central-directory record");
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const recordLength = 46 + nameLength + extraLength + commentLength;
    const next = offset + recordLength;
    if (next > centralDirectoryEnd) {
      throw new Error("ZIP mutation found a truncated central-directory record");
    }
    const name = decodeUtf8(source.subarray(offset + 46, offset + 46 + nameLength));
    if (name === targetPart) {
      const duplicate = source.slice(offset, next);
      const result = new Uint8Array(source.byteLength + duplicate.byteLength);
      result.set(source.subarray(0, centralDirectoryEnd), 0);
      result.set(duplicate, centralDirectoryEnd);
      result.set(source.subarray(centralDirectoryEnd), centralDirectoryEnd + duplicate.byteLength);
      const resultView = new DataView(result.buffer);
      const updatedEocdOffset = eocdOffset + duplicate.byteLength;
      resultView.setUint16(updatedEocdOffset + 8, entriesOnDisk + 1, true);
      resultView.setUint16(updatedEocdOffset + 10, entryCount + 1, true);
      resultView.setUint32(
        updatedEocdOffset + 12,
        centralDirectorySize + duplicate.byteLength,
        true,
      );
      return result;
    }
    offset = next;
  }
  throw new Error(`ZIP mutation could not find central directory entry for: ${targetPart}`);
};

/** Apply one deterministic malformed-package mutation to a DOCX ZIP. */
export const mutateZip = async (
  source: Uint8Array,
  mutation: ZipMutation,
  targetPart: string,
): Promise<Uint8Array> => {
  switch (mutation) {
    case "drop-part": {
      const zip = await JSZip.loadAsync(source);
      requirePart(zip, targetPart);
      zip.remove(targetPart);
      return generated(zip);
    }
    case "rename-part": {
      const zip = await JSZip.loadAsync(source);
      const part = requirePart(zip, targetPart);
      const renamed = `${targetPart}.renamed`;
      if (zip.file(renamed) !== null) {
        throw new Error(`ZIP mutation destination already exists: ${renamed}`);
      }
      zip.file(renamed, await part.async("uint8array"));
      zip.remove(targetPart);
      return generated(zip);
    }
    case "duplicate-part": {
      return duplicateCentralDirectoryEntry(source, targetPart);
    }
    case "truncate":
      return source.slice(0, Math.max(0, source.byteLength - 32));
    case "bad-crc":
      return badCentralDirectoryCrc(source, targetPart);
    case "wrong-content-type":
      return wrongContentType(source, targetPart);
    case "missing-relationship-target":
      return missingRelationshipTarget(source, targetPart);
  }
};
