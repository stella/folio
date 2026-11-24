import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { PreservedMarkup } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { createFolioAIEditSnapshot } from "./snapshot";

const XML =
  '<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

test(
  "opaque diagnostic identities depend only on emitted carrier order",
  () => {
    assertProperty(
      fc.property(
        fc.string({ maxLength: 80 }),
        fc.nat({ max: 5 }),
        fc.nat({ max: 5 }),
        (text, paragraphCount, markerCount) => {
          const paragraph = () =>
            schema.node("paragraph", null, [
              ...(text === "" ? [] : [schema.text(text)]),
              ...Array.from({ length: markerCount }, () =>
                schema.node("commentReference", { commentId: 7 }),
              ),
            ]);
          const diagnostics = (includeIgnoredChildren: boolean) => {
            const preserved = {
              children: [
                ...(includeIgnoredChildren
                  ? [{ index: 0, xml: "<w:proofErr/>" }]
                  : []),
                { index: 0, xml: XML },
                { index: 0, xml: XML },
              ],
            } satisfies PreservedMarkup;
            const doc = schema.node("doc", null, [
              ...(includeIgnoredChildren
                ? Array.from({ length: paragraphCount + 1 }, paragraph)
                : [schema.node("paragraph")]),
              schema.node("preservedBlock", { xml: XML }),
              schema.node("table", { _preserved: preserved }, [
                schema.node("tableRow", { _preserved: preserved }, [
                  schema.node("tableCell", null, [
                    paragraph(),
                    schema.node("preservedBlock", { xml: XML }),
                  ]),
                ]),
              ]),
              schema.node("preservedBlock", { xml: XML }),
            ]);
            return createFolioAIEditSnapshot(doc).blocks.filter(
              ({ kind }) => kind === "diagnostic",
            );
          };
          const before = diagnostics(false);
          expect(before).toHaveLength(7);
          expect(new Set(before.map(({ id }) => id)).size).toBe(before.length);
          expect(diagnostics(true)).toEqual(before);
        },
      ),
      {
        id: "opaque diagnostic identities depend only on emitted carrier order",
        numRuns: 50,
      },
    );
  },
  propertyTestTimeout(5_000),
);

test.each(["canonical"] as const)(
  "%s readers keep mixed opaque occurrences distinct through edits and save",
  async (session) => {
    const zip = await ZipArchive.read(await createEmptyDocx());
    zip.set(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:r><w:t>Before</w:t></w:r></w:p>
        <w:altChunk/><w:altChunk/>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>
          <w:p/><w:altChunk/><w:altChunk/>
        </w:tc></w:tr></w:tbl>
        <w:sectPr/>
      </w:body></w:document>`,
    );
    const source = await zip.generate({ compression: "store" });
    const prepared = (await prepareCanonicalDocxInput(source.buffer)).unwrap();
    const reviewer = await FolioDocxReviewer.fromBuffer(
      prepared.slice().buffer,
      {
        author: "Reviewer",
        session,
      },
    );
    const before = reviewer
      .getContent()
      .filter(({ kind }) => kind === "diagnostic");
    expect(before).toHaveLength(4);
    expect(new Set(before.map(({ id }) => id)).size).toBe(before.length);
    const paragraph = reviewer
      .getContent()
      .find(({ kind }) => kind === "paragraph");
    if (paragraph === undefined)
      panic("The carrier fixture lost its editable paragraph");
    const result = reviewer.applyOperations(
      [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: paragraph.id,
          find: "Before",
          replace: "Changed preceding content",
        },
      ],
      { mode: "direct", undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      {
        author: "Reviewer",
        session,
      },
    );
    for (const reader of [reviewer, reopened]) {
      expect(
        reader.getContent().filter(({ kind }) => kind === "diagnostic"),
      ).toEqual(before);
      const story = reader.readStory({ type: "main" });
      if (story === null) panic("The carrier fixture lost its main story");
      for (const { id, text } of before) {
        expect(story.text).toContain(`[${id}] ${text}`);
        expect(reader.getContentAsText({ annotated: true })).toContain(
          `[${id}] ${text}`,
        );
      }
    }
  },
);
