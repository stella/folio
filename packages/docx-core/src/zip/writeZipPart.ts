import type JSZip from "jszip";

// Match the deterministic timestamp used by the package fixtures.
const FIXED_ZIP_DATE_MS = Date.UTC(2000, 0, 1);
const UNIX_DIRECTORY_FLAG = 0x4000;
const DOS_DIRECTORY_FLAG = 0x10;

type WriteZipPartOptions = {
  zip: JSZip;
  path: string;
  data: Uint8Array | Promise<Uint8Array>;
  options?: Omit<JSZip.JSZipFileOptions, "createFolders" | "binary" | "base64">;
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
    data = new Uint8Array();
  }
  zip.file(path, data, {
    ...options,
    binary: true,
    base64: false,
    date: options.date ?? new Date(FIXED_ZIP_DATE_MS),
    createFolders: false,
  });
};
