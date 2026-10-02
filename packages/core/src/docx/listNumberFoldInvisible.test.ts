/**
 * A document that holds no `LISTNUM` field does not notice the fold.
 *
 * The digests below were taken with the reader, the projection and the save
 * as they stood before the fold kept its fields in the paragraph content:
 * the editor document each fixture projects to, and the `word/document.xml`
 * a save of that projection writes. They are compared byte for byte, so a
 * change here that reaches a paragraph it has no business with shows as a
 * digest that moved.
 *
 * `docx-editor-demo.docx` is the fixture with numbered paragraphs, the ones
 * the fold looks at. Its projection carries identifiers minted per process,
 * so only its length is pinned; its saved markup is pinned whole.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { createHash } from "node:crypto";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  normalizeFoldedListNumbers,
  unfoldPastedListNumberFields,
} from "../prosemirror/foldedListNumber";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

type Pinned = {
  path: string;
  /** SHA-256 of the projected document's JSON, or its length where ids are minted per process. */
  projection: { digest: string } | { length: number };
  /** SHA-256 of the saved `word/document.xml`. */
  saved: string;
};

const FIXTURES: readonly Pinned[] = [
  {
    path: "tests/visual/fixtures/sample.docx",
    projection: { digest: "2213484f1e631a4a2e3223ff8ec4fca38e5eb1ee4dea358fdca9b832aee3d04b" },
    saved: "4bb092166fcce424d39b691c4f395933baf70a3a769668a6b214b22ad6843fae",
  },
  {
    path: "packages/core/src/docx/__tests__/__fixtures__/corpus/step3-header-footer-fields.docx",
    projection: { digest: "a68195ab6c8e32a6df0e09ed04cc6131467dd2f5035fda94da7580149e7726aa" },
    saved: "264e9dd98b60a219429213a8ff7b19635762bf1d33f229fcbfc6918336cd48ad",
  },
  {
    path: "packages/core/src/docx/__tests__/__fixtures__/corpus/step3-footnotes.docx",
    projection: { digest: "db9ec37655bc5ed3c05a45c698f76a09f36fa5786d43bd843746ddcadc844c22" },
    saved: "5bb6efa1903f0a9b92920345f0781ccb1560f5ebf5275dba5d5d9f4a692888f4",
  },
  {
    path: "tests/visual/fixtures/docx-editor-demo.docx",
    projection: { length: 445_121 },
    saved: "6bac5af14f8a58ea6627768ea747a8fc8ec69ce3e3edae1c41d720de84f81dee",
  },
];

const sha256 = (text: string): string => createHash("sha256").update(text).digest("hex");

const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const xml = await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text");
  if (xml === undefined) {
    throw new Error("The package has no word/document.xml");
  }
  return xml;
};

describe("a document with no LISTNUM field", () => {
  for (const { path, projection, saved } of FIXTURES) {
    test(`${path} projects and saves as it did`, async () => {
      const buffer = await Bun.file(new URL(`../../../../${path}`, import.meta.url)).arrayBuffer();
      expect(/LISTNUM/iu.test(await documentXmlOf(buffer))).toBe(false);

      const parsed = await parseDocx(buffer, { preloadFonts: false, detectVariables: false });
      const doc = toProseDoc(parsed);
      const json = JSON.stringify(doc.toJSON());

      expect(json).not.toContain("foldedListNumber");
      if ("digest" in projection) {
        expect(sha256(json)).toBe(projection.digest);
      } else {
        expect(json).toHaveLength(projection.length);
      }

      const written = await repackDocx(fromProseDoc(doc, parsed), { updateModifiedDate: false });
      expect(sha256(await documentXmlOf(written))).toBe(saved);
    });

    test(`${path} gives the editor's own pass and a paste nothing to do`, async () => {
      const buffer = await Bun.file(new URL(`../../../../${path}`, import.meta.url)).arrayBuffer();
      const parsed = await parseDocx(buffer, { preloadFonts: false, detectVariables: false });
      const state = EditorState.create({ doc: toProseDoc(parsed) });

      expect(normalizeFoldedListNumbers(state)).toBeNull();
      const whole = state.doc.slice(0, state.doc.content.size);
      expect(unfoldPastedListNumberFields(whole)).toBe(whole);
    });
  }
});
