/**
 * Corpus invariant for the executable save boundary: a plain tracked insertion
 * needs no new package part, so saving it must preserve every ZIP entry name.
 * Directory entries are removed from inputs to expose JSZip's implicit folders.
 */

import { afterAll, describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { spawnSync } from "node:child_process";
import { readFile, readdir, writeFile } from "node:fs/promises";
import path from "node:path";

import { ensureParaIds, FolioDocxReviewer } from "@stll/folio-core/server";

import { makeTempDir } from "./__tests__/fixtures";
import { dataOf, ISOLATED_GIT_ENV } from "./__tests__/io";
import { fileVersionOf } from "./document";

const CORPUS = path.resolve(import.meta.dir, "../../core/src/docx/__tests__/__fixtures__/corpus");
const BIN = path.join(import.meta.dir, "bin.ts");
const FIXTURES = (await readdir(CORPUS)).filter((name) => name.endsWith(".docx")).toSorted();
const EXERCISED = new Set<string>();
const INSERTED_TEXT = "The parties confirm receipt of this copy.";
const PROCESS_TEST_TIMEOUT_MS = 120_000;

afterAll(() => {
  expect(FIXTURES.length).toBeGreaterThan(0);
  expect([...EXERCISED].toSorted()).toEqual(FIXTURES);
});

describe("folio suggest --in-place preserves corpus ZIP entry names", () => {
  test.each(FIXTURES)(
    "%s",
    async (name) => {
      const { dir, cleanup } = await makeTempDir();
      try {
        const inputZip = await JSZip.loadAsync(await readFile(path.join(CORPUS, name)));
        // JSZip.remove(directory) also removes its children; retain all parts.
        inputZip.files = Object.fromEntries(
          Object.entries(inputZip.files).filter(([, entry]) => !entry.dir),
        );
        const sourceNames = Object.keys(inputZip.files).toSorted();
        expect(sourceNames.some((entryName) => entryName.endsWith("/"))).toBe(false);
        const input = await inputZip.generateAsync({ type: "uint8array" });
        const { docx: source } = await ensureParaIds(input);
        expect(Object.keys((await JSZip.loadAsync(source)).files).toSorted()).toEqual(sourceNames);

        const reviewer = await FolioDocxReviewer.fromBuffer(source);
        // snapshot() includes blank paragraphs, so even the empty document is edited.
        const anchor = reviewer.snapshot().blocks.find(({ kind }) => kind !== "diagnostic");
        if (anchor === undefined) panic("Corpus fixture has no editable paragraph", { name });
        const file = path.join(dir, name);
        const operations = path.join(dir, "operations.json");
        await writeFile(file, source);
        await writeFile(
          operations,
          JSON.stringify([{ type: "insertAfterBlock", blockId: anchor.id, text: INSERTED_TEXT }]),
        );

        const run = spawnSync(
          process.execPath,
          [
            BIN,
            "suggest",
            file,
            "--input",
            `@${operations}`,
            "--in-place",
            "--allow-repack",
            "--expect-version",
            fileVersionOf(source),
            "--date",
            "2026-01-02T03:04:05Z",
          ],
          {
            cwd: dir,
            env: { ...ISOLATED_GIT_ENV, FOLIO_AUTHOR: "Corpus Reviewer" },
            encoding: "utf8",
            timeout: PROCESS_TEST_TIMEOUT_MS,
          },
        );
        expect(run.status, `${name}: ${run.stdout}\n${run.stderr}`).toBe(0);
        const receipt = dataOf(run.stdout);
        expect(receipt["status"]).toBe("committed");
        expect(receipt["fileVersion"]).not.toBe(fileVersionOf(source));

        const saved = new Uint8Array(await readFile(file));
        expect(Object.keys((await JSZip.loadAsync(saved)).files).toSorted()).toEqual(sourceNames);
        const reopened = await FolioDocxReviewer.fromBuffer(saved);
        expect(reopened.getContent().filter(({ text }) => text === INSERTED_TEXT)).toHaveLength(
          reviewer.getContent().filter(({ text }) => text === INSERTED_TEXT).length + 1,
        );
        EXERCISED.add(name);
      } finally {
        await cleanup();
      }
    },
    PROCESS_TEST_TIMEOUT_MS,
  );
});
