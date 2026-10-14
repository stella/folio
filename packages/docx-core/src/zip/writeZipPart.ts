import type JSZip from "jszip";

// Match the deterministic timestamp used by the package fixtures.
const FIXED_ZIP_DATE_MS = Date.UTC(2000, 0, 1);
const UNIX_DIRECTORY_FLAG = 0x4000;
const DOS_DIRECTORY_FLAG = 0x10;

// A part is complete text or bytes; streaming text owns separate decoder state.
type ZipPartContent = Exclude<Awaited<Parameters<JSZip["loadAsync"]>[0]>, NodeJS.ReadableStream>;

type WriteZipPartOptions = {
  zip: JSZip;
  path: string;
  data: ZipPartContent | Promise<ZipPartContent>;
  options?: Omit<JSZip.JSZipFileOptions, "createFolders">;
};

type EncodePartOptions = Pick<WriteZipPartOptions, "options"> & { data: ZipPartContent };

const encodePart = ({ data, options }: EncodePartOptions) => {
  if (typeof data !== "string") return data;
  if (options?.base64 === true) {
    return Uint8Array.from(atob(data), (character) => character.charCodeAt(0));
  }
  if (options?.binary === true) {
    return Uint8Array.from({ length: data.length }, (_, index) => data.charCodeAt(index));
  }
  return new TextEncoder().encode(data);
};

type WriteBytesOptions = Omit<WriteZipPartOptions, "data"> & {
  data: Exclude<ZipPartContent, string> | Promise<Exclude<ZipPartContent, string>>;
};

const writeBytes = ({ zip, path, data, options }: WriteBytesOptions): void => {
  zip.file(path, data, {
    ...options,
    binary: true,
    base64: false,
    date: options?.date ?? new Date(FIXED_ZIP_DATE_MS),
    createFolders: false,
  });
};

/** Write exactly one entry, without synthesizing parent directory entries. */
export const writeZipPart = ({ zip, path, data, options = {} }: WriteZipPartOptions): void => {
  const unixPermissions =
    typeof options.unixPermissions === "string"
      ? Number.parseInt(options.unixPermissions, 8)
      : (options.unixPermissions ?? 0);
  if (
    options.dir === true ||
    (unixPermissions & UNIX_DIRECTORY_FLAG) !== 0 ||
    ((options.dosPermissions ?? 0) & DOS_DIRECTORY_FLAG) !== 0
  ) {
    writeBytes({ zip, path, data: new Uint8Array(), options });
    return;
  }
  // JSZip chunks strings before UTF-8 encoding, which can divide a surrogate pair.
  // Encode text as a whole so every later archive chunk contains only bytes.
  const content =
    typeof data === "object" && data !== null && "then" in data
      ? data.then((resolved) => encodePart({ data: resolved, options }))
      : encodePart({ data, options });
  writeBytes({ zip, path, data: content, options });
};
