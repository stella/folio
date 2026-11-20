import { expect, test } from "bun:test";
import { readFile } from "node:fs/promises";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
} from "@stll/folio-core/server";

import { makeTempDir } from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import { packageTexts, writePackage } from "./__tests__/package-fixtures";
import { runFolioCli } from "./cli";

const header = { type: "header", relationshipId: "rIdHeader" } as const;
const footer = { type: "footer", relationshipId: "rIdFooter" } as const;
const paragraph = (id: string) =>
  `<w:p w14:paraId="${id}"><w:r><w:t>Draft</w:t></w:r></w:p>`;

test.each([header, footer])(
  "suggest writes only the selected $type",
  async (story) => {
    const { dir, cleanup } = await makeTempDir();
    try {
      const file = await writePackage(dir, "draft.docx", {
        body: paragraph("10000001"),
        headers: { [header.relationshipId]: paragraph("10000002") },
        footers: { [footer.relationshipId]: paragraph("10000003") },
        sectPr:
          '<w:headerReference w:type="default" r:id="rIdHeader"/>' +
          '<w:footerReference w:type="default" r:id="rIdFooter"/>',
      });
      const run = async (command: string, ...flags: string[]) => {
        const captured = captureIo();
        const exit = await runFolioCli([command, file, ...flags], captured.io);
        return { exit, stdout: captured.stdout() };
      };
      const discovered = await run("stories");
      expect(discovered.exit, discovered.stdout).toBe(0);
      expect(dataOf(discovered.stdout)["result"]).toEqual(
        expect.arrayContaining([expect.objectContaining({ handle: story })]),
      );
      const read = await run("story", "--handle", JSON.stringify(story));
      expect(read.exit, read.stdout).toBe(0);
      const blockId = story.type === "header" ? "10000002" : "10000003";
      expect(dataOf(read.stdout)["result"]).toMatchObject({
        handle: story,
        blocks: [{ blockId, text: "Draft", blockTextHash: expect.any(String) }],
      });
      const version = dataOf(read.stdout)["fileVersion"];
      if (typeof version !== "string")
        throw new Error("Expected a file version");
      const before = await readFile(file);
      const suggest = (operations: unknown, handle: unknown = story) =>
        run(
          "suggest",
          "--story",
          JSON.stringify(handle),
          "--operations",
          JSON.stringify(operations),
          "--in-place",
          "--expect-version",
          version,
        );
      const replaceAll = {
        type: "replaceAll",
        find: "Draft",
        replace: "Final",
      };
      const mixed = await suggest([
        replaceAll,
        {
          type: "replaceInBlock",
          blockId: "10000001",
          find: "Draft",
          replace: "Wrong",
        },
      ]);
      expect(mixed.exit).not.toBe(0);
      expect(await readFile(file)).toEqual(before);
      const missing = await suggest([replaceAll], {
        type: story.type,
        relationshipId: "rIdMissing",
      });
      expect(envelopeOf(missing.stdout)["error"]).toMatchObject({
        code: "not_found",
      });
      expect(await readFile(file)).toEqual(before);
      const invalid = await suggest([replaceAll], { type: story.type });
      expect(envelopeOf(invalid.stdout)["error"]).toMatchObject({
        code: "invalid_input",
      });
      expect(await readFile(file)).toEqual(before);

      const applied = await suggest([replaceAll]);
      expect(applied.exit, applied.stdout).toBe(0);
      const after = await readFile(file);
      const originalParts = await packageTexts(before);
      const savedParts = await packageTexts(after);
      const other = story.type === "header" ? footer : header;
      const part = (handle: typeof header | typeof footer) =>
        `word/${handle.type}-${handle.relationshipId}.xml`;
      expect(savedParts.get("word/document.xml")).toBe(
        originalParts.get("word/document.xml"),
      );
      expect(savedParts.get(part(other))).toBe(originalParts.get(part(other)));
      expect(savedParts.get(part(story))).toMatch(/<w:ins\b/);
      expect(savedParts.get(part(story))).toMatch(/<w:del\b/);
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await Bun.file(file).arrayBuffer(),
      );
      expect(reviewer.readStory(story)?.text).toBe("Final");
      reviewer.rejectAll();
      expect(reviewer.readStory(story)?.text).toBe("Draft");
    } finally {
      await cleanup();
    }
  },
);

for (const session of ["legacy", "canonical"] as const)
  test.each([header, footer])(
    `a $type edit saves without a repack on the ${session} engine`,
    async (story) => {
      const { dir, cleanup } = await makeTempDir();
      try {
        const file = await writePackage(dir, "draft.docx", {
          body: paragraph("10000001"),
          headers: { [header.relationshipId]: paragraph("10000002") },
          footers: { [footer.relationshipId]: paragraph("10000003") },
          sectPr:
            '<w:headerReference w:type="default" r:id="rIdHeader"/>' +
            '<w:footerReference w:type="default" r:id="rIdFooter"/>',
        });
        const before = await readFile(file);
        const reviewer = await FolioDocxReviewer.fromBuffer(
          before.slice().buffer,
          { session, author: "Reviewer" },
        );
        const result = reviewer.applyDocumentOperationsToStory({
          story,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [
              {
                id: "story-edit",
                type: "replaceInBlock",
                blockId: story.type === "header" ? "10000002" : "10000003",
                find: "Draft",
                replace: "Final",
              },
            ],
          },
          undefinedReferences: "refuse",
        });
        expect(result.skipped).toEqual([]);
        const saved = await reviewer.save({ repack: "refuse" });
        if (saved.type !== "selective")
          throw new Error(`Expected a selective save, got ${saved.type}`);
        const originalParts = await packageTexts(before);
        const savedParts = await packageTexts(new Uint8Array(saved.buffer));
        const other = story.type === "header" ? footer : header;
        const part = (handle: typeof header | typeof footer) =>
          `word/${handle.type}-${handle.relationshipId}.xml`;
        expect(savedParts.get("word/document.xml")).toBe(
          originalParts.get("word/document.xml"),
        );
        expect(savedParts.get(part(other))).toBe(
          originalParts.get(part(other)),
        );
        const reopened = await FolioDocxReviewer.fromBuffer(saved.buffer, {
          session,
        });
        expect(reopened.readStory(story)?.text).toBe("Final");
        reopened.rejectAll();
        expect(reopened.readStory(story)?.text).toBe("Draft");
      } finally {
        await cleanup();
      }
    },
  );
