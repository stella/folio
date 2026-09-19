/**
 * No folio authoring path may write a style reference the package it produces
 * does not define.
 *
 * The hardcoded English built-in id is the recurring form of this defect
 * (`Heading${n}`, `TOCHeading`, `Quote`, `CommentReference`), and it is
 * systematic rather than input-dependent: a writer that names a style it does
 * not define does so for every document it writes. Holding every authoring
 * path to one check is what keeps the class closed, rather than fixing the
 * instances as they are noticed.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { compileLegalSourceToDocument } from "@stll/docx-core";

import { fromMarkdown } from "../markdown/fromMarkdown";
import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";
import {
  createTableOfContentsField,
  HEADING_LEVELS,
  heading,
  paragraph,
  table,
} from "./server/build";
import { danglingStyleReferences } from "./styleReferenceResolution";

/** Every part a style reference can appear in. */
const REFERENCING_PARTS =
  /^word\/(document|styles|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/u;

const danglingIn = async (document: Document): Promise<string[]> => {
  const zip = await JSZip.loadAsync(await createDocx(document));
  const parts = new Map<string, string>();
  for (const name of Object.keys(zip.files)) {
    if (REFERENCING_PARTS.test(name)) {
      parts.set(name, await zip.file(name)!.async("string"));
    }
  }
  return danglingStyleReferences(parts.get("word/styles.xml"), parts).map(
    ({ styleId, part }) => `${part}: ${styleId}`,
  );
};

describe("a document folio authors resolves every style it references", () => {
  test("an empty document, from each built-in style set", async () => {
    expect(await danglingIn(createEmptyDocument())).toEqual([]);
    expect(
      await danglingIn(createEmptyDocument({ preset: createStellaStyleDocumentPreset() })),
    ).toEqual([]);
  });

  test("markdown import, over arbitrary markdown", async () => {
    // Heading depth, blockquotes and lists are the constructs that reach for a
    // built-in style id, so the generator has to be able to produce all three.
    const block = fc.oneof(
      fc
        .tuple(fc.integer({ min: 1, max: 6 }), fc.constantFrom("Alpha", "Beta", "Gamma"))
        .map(([depth, text]) => `${"#".repeat(depth)} ${text}`),
      fc.constantFrom("> quoted", "- item\n- item", "1. one\n2. two", "plain text", "`code`"),
    );
    await fc.assert(
      fc.asyncProperty(fc.array(block, { minLength: 1, maxLength: 8 }), async (blocks) => {
        expect(await danglingIn(fromMarkdown(blocks.join("\n\n")))).toEqual([]);
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("the report builder, for every heading level it accepts", async () => {
    const document = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    document.package.document.content = [
      createTableOfContentsField(),
      ...HEADING_LEVELS.map((level) => heading({ text: `Level ${level}`, level })),
      paragraph("Body."),
      table({ header: ["A", "B"], rows: [["a", "b"]] }),
    ];
    expect(await danglingIn(document)).toEqual([]);
  });

  test("the legal-source compiler", async () => {
    const compiled = compileLegalSourceToDocument(
      [
        "# Sale Agreement",
        "",
        "## 1. Definitions",
        "",
        "Agreement means this contract.",
        "",
        "## 2. Term",
        "",
        "- first",
        "- second",
      ].join("\n"),
    );
    if (compiled.status !== "ok") {
      throw new Error(`legal source did not compile: ${compiled.status}`);
    }
    expect(await danglingIn(compiled.document)).toEqual([]);
  });

  test("the report builder over the generic style set, not only the stella one", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [
      createTableOfContentsField(),
      ...HEADING_LEVELS.map((level) => heading({ text: `Level ${level}`, level })),
      paragraph("Body."),
      table({ header: ["A", "B"], rows: [["a", "b"]] }),
    ];
    expect(await danglingIn(document)).toEqual([]);
  });

  test("a document carrying a comment and a footnote", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [paragraph("Body.")];
    document.package.comments = [
      {
        id: 1,
        author: "A",
        initials: "A",
        date: "2024-01-01T00:00:00Z",
        content: [paragraph("A remark.")],
      },
    ];
    document.package.footnotes = [{ type: "footnote", id: 1, content: [paragraph("A note.")] }];
    expect(await danglingIn(document)).toEqual([]);
  });

  test("the same, for a document that carries no style table of its own", async () => {
    // The seed package defines `docDefaults` and `Normal` and nothing else, so
    // this is the one authoring path the model's style table cannot repair.
    const document = createEmptyDocument();
    document.package.styles = undefined;
    document.package.document.content = [paragraph("Body.")];
    document.package.comments = [
      {
        id: 1,
        author: "A",
        initials: "A",
        date: "2024-01-01T00:00:00Z",
        content: [paragraph("A remark.")],
      },
    ];
    document.package.footnotes = [{ type: "footnote", id: 1, content: [paragraph("A note.")] }];
    expect(await danglingIn(document)).toEqual([]);
  });
});
