/** Read a JSON data asset in browsers and supported Node runtimes. */
export const readRuntimeJsonAsset = async (url: URL): Promise<unknown> => {
  if (url.protocol !== "file:") {
    const response = await fetch(url);
    if (!response.ok)
      throw new Error(`${url.href} answered ${response.status}`);
    return response.json();
  }

  // Node's fetch does not read files. Reach the built-in module indirectly so
  // browser bundlers do not see a `node:` import.
  const runtimeProcess: unknown = Reflect.get(globalThis, "process");
  if (
    typeof runtimeProcess !== "object" ||
    runtimeProcess === null ||
    !("getBuiltinModule" in runtimeProcess) ||
    typeof runtimeProcess.getBuiltinModule !== "function"
  ) {
    throw new Error(`This runtime cannot read ${url.href}`);
  }
  const fs: unknown = runtimeProcess.getBuiltinModule("node:fs/promises");
  if (
    typeof fs !== "object" ||
    fs === null ||
    !("readFile" in fs) ||
    typeof fs.readFile !== "function"
  ) {
    throw new Error(`This runtime cannot read ${url.href}`);
  }
  return JSON.parse(await fs.readFile(url, "utf8"));
};
