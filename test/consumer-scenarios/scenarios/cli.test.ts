/**
 * The `folio` command line on the same synthetic documents: read, suggest,
 * comment, accept and save, each `--in-place` against the `fileVersion` of
 * the read before it; after each, core reopens the file and every reader
 * agrees. A stale version is refused and writes nothing.
 */

import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, test } from "node:test";

import { FIXTURES, type FixtureName, openReviewer } from "../support/documents.ts";
import { assertHealthy, assertReadersAgree } from "../support/invariants.ts";
import { folio } from "../support/cli.ts";

type ReadBlock = { blockId: string; kind: string; text: string; blockTextHash: string };

const CLI_FIXTURES: readonly FixtureName[] = [
  "plain",
  "lists",
  "styleNumbered",
  "tables",
  "notes",
  "comments",
  "trackedChanges",
];

const read = async (file: string, cwd: string) => {
  const run = await folio(["read", file], cwd);
  assert.equal(run.code, 0, `folio read failed: ${JSON.stringify(run.json)}`);
  const data = run.json.data as { fileVersion: string; result: { blocks: ReadBlock[] } };
  return { version: data.fileVersion, blocks: data.result.blocks };
};

const expectCommitted = (run: Awaited<ReturnType<typeof folio>>, what: string): void => {
  assert.equal(run.code, 0, `${what} failed (exit ${run.code}): ${JSON.stringify(run.json)}`);
  assert.equal(run.json.ok, true, `${what} did not commit`);
};

const reopenAndCheck = async (file: string, context: string) => {
  const bytes = new Uint8Array(await readFile(file));
  await assertReadersAgree(bytes, context);
  return assertHealthy(await openReviewer(bytes), `${context} (resaved by core)`);
};

describe("folio CLI", () => {
  for (const name of CLI_FIXTURES) {
    test(`${name}: read → suggest → comment → accept → save, in place`, async () => {
      const cwd = await mkdtemp(path.join(tmpdir(), "folio-cli-scenario-"));
      try {
        const file = path.join(cwd, "document.docx");
        await writeFile(file, await FIXTURES[name]());

        const first = await read(file, cwd);
        const target = first.blocks.find(
          (block) => block.kind === "paragraph" && /\w{4,}/u.test(block.text),
        );
        assert.ok(target, "no paragraph to edit");
        const word = /\w{4,}/u.exec(target.text)?.[0] ?? "";
        const operations = [
          { type: "replaceInBlock", blockId: target.blockId, find: word, replace: "amended" },
          {
            type: "insertAfterBlock",
            blockId: target.blockId,
            text: "A sentence the CLI added.",
            precondition: { blockTextHash: target.blockTextHash },
          },
        ];
        const suggested = await folio(
          [
            "suggest",
            file,
            "--input",
            JSON.stringify({ operations }),
            "--in-place",
            "--allow-repack",
            "--expect-version",
            first.version,
          ],
          cwd,
        );
        expectCommitted(suggested, "folio suggest");
        const { reopened } = await reopenAndCheck(file, `${name} after folio suggest`);
        assert.ok(
          reopened.getChanges().some((change) => change.author === "CLI Scenario"),
          "the suggestion is not a tracked change by the CLI author",
        );

        // A write against the version before the suggestion is refused and
        // leaves the file alone.
        const bytesBefore = await readFile(file);
        const stale = await folio(
          ["comment", file, "--block-id", target.blockId, "--text", "Stale.", "--in-place"].concat([
            "--expect-version",
            first.version,
          ]),
          cwd,
        );
        assert.equal(
          stale.code,
          10,
          `a stale write was not refused: ${JSON.stringify(stale.json)}`,
        );
        assert.deepEqual(await readFile(file), bytesBefore, "a refused write changed the file");

        const second = await read(file, cwd);
        const commented = await folio(
          [
            "comment",
            file,
            "--block-id",
            target.blockId,
            "--text",
            "Please confirm.",
            "--in-place",
            "--expect-version",
            second.version,
          ],
          cwd,
        );
        expectCommitted(commented, "folio comment");
        await reopenAndCheck(file, `${name} after folio comment`);

        const third = await read(file, cwd);
        const accepted = await folio(
          [
            "accept",
            file,
            "--all",
            "--in-place",
            "--allow-repack",
            "--expect-version",
            third.version,
          ],
          cwd,
        );
        expectCommitted(accepted, "folio accept --all");
        const { reopened: acceptedReviewer } = await reopenAndCheck(
          file,
          `${name} after folio accept`,
        );
        assert.deepEqual(acceptedReviewer.getChanges(), [], "accept --all left changes behind");
        const texts = acceptedReviewer.getContent().map((block) => block.text);
        assert.ok(texts.includes("A sentence the CLI added."), "the accepted insertion is missing");

        // An editor's whole-package save through the lease.
        const fourth = await read(file, cwd);
        const editor = await openReviewer(new Uint8Array(await readFile(file)));
        const last = editor.getContent().at(-1);
        assert.ok(last);
        editor.applyDocumentOperations({
          version: 1,
          mode: "direct",
          operations: [
            { id: "1", type: "insertAfterBlock", blockId: last.id, text: "Saved by an editor." },
          ],
        } as never);
        const editorFile = path.join(cwd, "editor-save.docx");
        await writeFile(editorFile, new Uint8Array(await editor.toBuffer()));
        const saved = await folio(
          ["save", file, "--from", editorFile, "--expect-version", fourth.version],
          cwd,
        );
        expectCommitted(saved, "folio save");
        const { reopened: final } = await reopenAndCheck(file, `${name} after folio save`);
        assert.equal(final.getContent().at(-1)?.text, "Saved by an editor.");
      } finally {
        await rm(cwd, { recursive: true, force: true });
      }
    });
  }
});
