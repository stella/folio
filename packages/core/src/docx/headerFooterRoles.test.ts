/**
 * A header or footer part's role (`default`, `first`, `even`) is stated by
 * the section reference that names it (ECMA-376 Part 1 §17.10.5), not by the
 * part. The package parser reads the parts from the document's relationships,
 * so it has to take each role from the sections; it used to read every part
 * as a default one, and readers labelled first-page and even-page headers as
 * default headers.
 */

import { describe, expect, test } from "bun:test";

import { buildSectionsDocx, SECTIONS_FIXTURE_PARTS } from "../ai-edits/__fixtures__/sections";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";

const rolesOf = (parts: Map<string, { hdrFtrType: string }> | undefined) =>
  Object.fromEntries([...(parts ?? [])].map(([id, part]) => [id, part.hdrFtrType]));

describe("header and footer roles", () => {
  test("each part takes the role its first section reference gives it", async () => {
    const document = await parseDocx(await buildSectionsDocx());

    expect(rolesOf(document.package.headers)).toEqual({
      rIdHeaderOneDefault: "default",
      rIdHeaderOneFirst: "first",
      rIdHeaderTwoDefault: "default",
      rIdHeaderTwoEven: "even",
      rIdHeaderThreeDefault: "default",
    });
    expect(rolesOf(document.package.footers)).toEqual({ rIdFooterOneDefault: "default" });
  });

  test("the roles survive a repack and a reparse", async () => {
    const source = await buildSectionsDocx();
    const document = await parseDocx(source);
    const reparsed = await parseDocx(await repackDocx({ ...document, originalBuffer: source }));

    expect(rolesOf(reparsed.package.headers)).toEqual(rolesOf(document.package.headers));
  });

  test("readers label each part with its roles and the sections that show it", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildSectionsDocx());
    const text = (id: keyof typeof SECTIONS_FIXTURE_PARTS) => SECTIONS_FIXTURE_PARTS[id].text;

    expect(reviewer.getNotesAsText().split("\n")).toEqual([
      `[header default (section 1)] ${text("rIdHeaderOneDefault")}`,
      `[header first (sections 1, 3)] ${text("rIdHeaderOneFirst")}`,
      `[header default (section 2)] ${text("rIdHeaderTwoDefault")}`,
      `[header even (section 2)] ${text("rIdHeaderTwoEven")}`,
      `[header default (section 3)] ${text("rIdHeaderThreeDefault")}`,
      `[footer default (section 1)] ${text("rIdFooterOneDefault")}`,
    ]);
  });

  test("a single section's first and even headers are labelled without section numbers", async () => {
    const paragraph = (text: string): Paragraph => ({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text }] }],
    });
    const model = fromMarkdown("Body.");
    model.package.headers = new Map([
      ["rIdFirst", { type: "header", hdrFtrType: "first", content: [paragraph("First page.")] }],
      ["rIdEven", { type: "header", hdrFtrType: "even", content: [paragraph("Even page.")] }],
    ]);
    model.package.document.finalSectionProperties = {
      ...model.package.document.finalSectionProperties,
      titlePg: true,
      evenAndOddHeaders: true,
      headerReferences: [
        { type: "first", rId: "rIdFirst" },
        { type: "even", rId: "rIdEven" },
      ],
    };
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(model));

    expect(rolesOf(reviewer.toDocument().package.headers)).toEqual({
      rIdFirst: "first",
      rIdEven: "even",
    });
    expect(reviewer.getNotesAsText().split("\n")).toEqual([
      "[header first] First page.",
      "[header even] Even page.",
    ]);
  });
});
