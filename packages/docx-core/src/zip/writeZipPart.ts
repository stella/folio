import type JSZip from "jszip";

// Match the deterministic timestamp used by the package fixtures.
const FIXED_ZIP_DATE_MS = Date.UTC(2000, 0, 1);

type WriteZipPartOptions = {
  zip: JSZip;
  path: string;
  data: Parameters<JSZip["loadAsync"]>[0];
  options?: Omit<JSZip.JSZipFileOptions, "createFolders">;
};

/** Write exactly one entry, without synthesizing parent directory entries. */
export const writeZipPart = ({ zip, path, data, options }: WriteZipPartOptions): void => {
  zip.file(path, data, {
    ...options,
    date: options?.date ?? new Date(FIXED_ZIP_DATE_MS),
    createFolders: false,
  });
};
