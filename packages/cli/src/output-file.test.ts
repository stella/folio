import {
  afterEach,
  beforeEach,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import { mkdir, readdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../../../test/subprocess";
import { makeTempDir } from "./__tests__/fixtures";
import { writeOutputFile } from "./output-file";

setDefaultTimeout(subprocessTestTimeout());

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterEach(async () => {
  await cleanup();
});

const identityOf = async (filePath: string) => {
  const info = await stat(filePath);
  return { dev: info.dev, ino: info.ino };
};

describe("writeOutputFile", () => {
  test("writes beside the request and replaces only with overwrite", async () => {
    const input = path.join(dir, "contract.docx");
    await writeFile(input, "docx");
    const target = path.join(dir, "contract.pdf");
    const bytes = new TextEncoder().encode("pdf");

    const first = await writeOutputFile({
      target,
      bytes,
      overwrite: false,
      input: await identityOf(input),
    });
    expect(first.isOk()).toBe(true);
    expect(await readFile(target, "utf8")).toBe("pdf");

    const refused = await writeOutputFile({
      target,
      bytes,
      overwrite: false,
      input: await identityOf(input),
    });
    expect(refused.isErr() && refused.error.code).toBe("destination_exists");

    const replaced = await writeOutputFile({
      target,
      bytes: new TextEncoder().encode("pdf 2"),
      overwrite: true,
      input: await identityOf(input),
    });
    expect(replaced.isOk()).toBe(true);
    expect(await readFile(target, "utf8")).toBe("pdf 2");
    expect((await readdir(dir)).toSorted()).toEqual([
      "contract.docx",
      "contract.pdf",
    ]);
  });

  test("an output directory moved after it was resolved receives nothing, nor does its replacement", async () => {
    const output = path.join(dir, "out");
    const elsewhere = path.join(dir, "elsewhere");
    await mkdir(output);
    await mkdir(elsewhere);
    const input = path.join(dir, "contract.docx");
    await writeFile(input, "docx");
    const modulePath = path.join(import.meta.dir, "output-file.ts");
    // The directory is swapped for a link to another one between the
    // writer's resolution and its first open.
    const source = `
      import { mock } from "bun:test";
      import { rename, symlink } from "node:fs/promises";
      import * as promises from "node:fs/promises";
      const realOpen = promises.open;
      let swapped = false;
      mock.module("node:fs/promises", () => ({
        ...promises,
        open: async (...args) => {
          if (!swapped) {
            swapped = true;
            await rename(${JSON.stringify(output)}, ${JSON.stringify(`${output}-moved`)});
            await symlink(${JSON.stringify(elsewhere)}, ${JSON.stringify(output)});
          }
          return await realOpen(...args);
        },
      }));
      const { writeOutputFile } = await import(${JSON.stringify(modulePath)});
      const result = await writeOutputFile({
        target: ${JSON.stringify(path.join(output, "contract.pdf"))},
        bytes: new TextEncoder().encode("pdf"),
        overwrite: false,
        input: ${JSON.stringify(await identityOf(input))},
      });
      console.log(JSON.stringify(result.isErr() ? result.error.code : "ok"));
    `;
    const result = expectExited(runChild([process.execPath, "--eval", source]));
    expect(result.stderr).toBe("");
    expect(JSON.parse(result.stdout)).toBe("unsafe_path");
    expect(await readdir(elsewhere)).toEqual([]);
    expect(await readdir(`${output}-moved`)).toEqual([]);
  });
});
