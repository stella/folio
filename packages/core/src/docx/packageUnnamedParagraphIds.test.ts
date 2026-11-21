/**
 * The quick check that lets a package naming every paragraph skip the id
 * assignment never skips one the assignment would give an id.
 */
import { Result } from "better-result";
import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  assignPackageParaIds,
  EnsureParaIdsError,
  packageStoryParts,
} from "./ensureParaIds";
import { unnamedPackageParagraphIds } from "./packageUnnamedParagraphIds";
import { unzipDocx } from "./unzip";

const CORPUS = path.join(import.meta.dir, "__tests__/__fixtures__/corpus");

const named = (
  ids: ReadonlyMap<string, { byOrdinal: ReadonlyMap<number, string> }>,
) =>
  [...ids].flatMap(([part, { byOrdinal }]) =>
    [...byOrdinal].map(([ordinal, id]) => `${part}#${ordinal}=${id}`),
  );

describe("unnamed package paragraph ids", () => {
  test.each(readdirSync(CORPUS).filter((name) => name.endsWith(".docx")))(
    "names exactly what the full assignment names (%s)",
    async (name) => {
      const raw = await unzipDocx(
        readFileSync(path.join(CORPUS, name)).buffer,
        {
          extractAllXml: false,
        },
      );
      const documentPartName = [...raw.allXml.keys()].find(
        (part) => part.toLowerCase() === "word/document.xml",
      );
      if (documentPartName === undefined) return;
      const { isTargetPart } = packageStoryParts({
        documentPartName,
        documentRels: raw.documentRels,
        entries: [...raw.allXml.keys()],
      });
      const assigned = Result.try({
        try: () =>
          assignPackageParaIds({
            partTexts: raw.allXml,
            documentPartName,
            isTargetPart,
          }),
        catch: (cause) => cause,
      });
      if (assigned.isErr() && !(assigned.error instanceof EnsureParaIdsError))
        throw assigned.error;
      const full = assigned.isOk() ? named(assigned.value) : [];
      expect(
        named(unnamedPackageParagraphIds(raw.allXml, raw.documentRels)),
      ).toEqual(full);
    },
  );
});
