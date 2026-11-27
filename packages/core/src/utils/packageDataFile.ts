/**
 * Data files the package ships beside its modules (hyphenation dictionaries,
 * locale catalogs) rather than as modules, so no bundle carries them. Each is
 * addressed by a literal `new URL(..., import.meta.url)` in its caller, which
 * bundlers that copy referenced files emit as an asset.
 */

/**
 * `<baseUrl>/<file>`: where a host keeps the package's data files of one
 * kind (a CDN path of the same version, say). `baseUrl` is absolute.
 */
export const dataFileUnder = (baseUrl: string | URL, file: string): URL => {
  const url = new URL(baseUrl);
  if (!url.pathname.endsWith("/")) {
    url.pathname += "/";
  }
  url.pathname += file;
  return url;
};

/**
 * The file reader of a runtime with Node's built-in modules, reached through
 * `process.getBuiltinModule` so that no bundler sees a `node:` import.
 */
const builtinReadFile = ():
  ((url: URL, encoding: "utf8") => Promise<string>) | undefined => {
  const runtimeProcess: unknown = Reflect.get(globalThis, "process");
  if (
    typeof runtimeProcess !== "object" ||
    runtimeProcess === null ||
    !("getBuiltinModule" in runtimeProcess) ||
    typeof runtimeProcess.getBuiltinModule !== "function"
  ) {
    return undefined;
  }
  const fs: unknown = runtimeProcess.getBuiltinModule("node:fs/promises");
  if (
    typeof fs !== "object" ||
    fs === null ||
    !("readFile" in fs) ||
    typeof fs.readFile !== "function"
  ) {
    return undefined;
  }
  const { readFile } = fs;
  return (url, encoding) => readFile(url, encoding);
};

/**
 * The parsed JSON at `url`. Node's fetch does not read files, so a package
 * installed on disk is read through the runtime's file system; a runtime
 * without `process.getBuiltinModule` (Node before 22.3) is handed the data
 * instead. Throws when the file cannot be read or is not JSON.
 */
export const readDataFile = async (url: URL): Promise<unknown> => {
  if (url.protocol === "file:") {
    const readFile = builtinReadFile();
    if (readFile === undefined) {
      throw new Error(`This runtime cannot read ${url.href}`);
    }
    return JSON.parse(await readFile(url, "utf8"));
  }
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`${url.href} answered ${response.status}`);
  }
  return response.json();
};
