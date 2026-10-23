import type JSZip from "jszip";

import { writeZipEntry } from "@stll/utf8-boundaries";

// Match the deterministic timestamp used by the package fixtures.
const FIXED_ZIP_DATE_MS = Date.UTC(2000, 0, 1);
const UNIX_DIRECTORY_FLAG = 0x4000;
const DOS_DIRECTORY_FLAG = 0x10;

type WriteZipPartOptions = {
  zip: JSZip;
  path: string;
  data: Parameters<JSZip["loadAsync"]>[0];
  options?: Omit<JSZip.JSZipFileOptions, "createFolders">;
};

/** Write exactly one entry, without synthesizing parent directory entries. */
export const writeZipPart = ({
  zip,
  path,
  data,
  options = {},
}: WriteZipPartOptions): void => {
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
  writeZipEntry(zip, path, data, {
    ...options,
    date: options.date ?? new Date(FIXED_ZIP_DATE_MS),
    createFolders: false,
  });
};
