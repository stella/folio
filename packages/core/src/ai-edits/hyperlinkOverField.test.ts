/**
 * A link applied over a selection that holds a field keeps the field and its
 * result: the runs around it are linked, the field stands between them.
 */

import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { shapeArrayBuffer } from "../__tests__/documentShapes";
import {
  createHarnessState,
  harnessManager,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { FolioDocxReviewer } from "./headless";

describe("setHyperlink over a paragraph with fields", () => {
  const link = async (mode: "editing" | "suggesting") => {
    const base = await parseShapeDocument(
      new Uint8Array(await shapeArrayBuffer("fields-links-bookmarks")),
    );
    const state = createHarnessState(base, mode);
    let from = -1;
    let to = -1;
    state.doc.descendants((node, position) => {
      if (node.isTextblock && node.textContent.startsWith("Page")) {
        from = position + 1;
        to = from + node.content.size;
      }
      return true;
    });
    const selected = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to)),
    );
    const factory = harnessManager().getCommand("setHyperlink");
    if (!factory) throw new Error("Missing command: setHyperlink");
    let next = selected;
    expect(
      factory("https://example.org/")(selected, (tr: Transaction) => {
        next = selected.apply(tr);
      }),
    ).toBe(true);
    const { bytes } = await saveHarnessState(next, base);
    return bytes.slice().buffer;
  };

  test.each(["editing", "suggesting"] as const)(
    "keeps the fields and their results (%s)",
    async (mode) => {
      const saved = await link(mode);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(reopened.getContent().map(({ text }) => text)).toContain(
        "Page 1 of 1",
      );
      const xml =
        (await (
          await ZipArchive.read(saved)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";
      expect(xml).toContain("w:fldSimple");
      expect(xml).toContain("fldChar");
      expect(xml).toContain("<w:hyperlink");
    },
  );
});
