import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";

import type { Document, HeaderFooter } from "../types/document";
import {
  createHeaderFooterEditorManager,
  enumerateDocumentHeaderFooterParts,
  enumerateHeaderFooterParts,
} from "./headerFooterEditorManager";

const headerFooter = (type: HeaderFooter["type"]): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [],
});

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe("header/footer editor slot identity", () => {
  test("uses one relationship-id slot for every shared part", () => {
    const headers = new Map([["rIdShared", headerFooter("header")]]);
    const footers = new Map([
      ["rIdShared", headerFooter("footer")],
      ["rIdFooter", headerFooter("footer")],
    ]);

    expect(enumerateHeaderFooterParts({ headers, footers })).toEqual([
      { kind: "header", rId: "rIdShared" },
      { kind: "footer", rId: "rIdFooter" },
    ]);
  });

  test("enumerates package parts without depending on section count", () => {
    const document: Document = {
      package: {
        document: {
          content: [],
          sections: [
            {
              content: [],
              properties: {
                headerReferences: [{ type: "default", rId: "rIdH" }],
              },
            },
            {
              content: [],
              properties: {
                headerReferences: [{ type: "default", rId: "rIdH" }],
              },
            },
          ],
        },
        headers: new Map([["rIdH", headerFooter("header")]]),
      },
    };

    expect(enumerateDocumentHeaderFooterParts(document)).toEqual([
      { kind: "header", rId: "rIdH" },
    ]);
  });
});

test.each(["header", "footer"] as const)(
  "persistent %s views sync read-only mode and reject document changes",
  (kind) => {
    const host = globalThis.document.createElement("div");
    globalThis.document.body.append(host);
    const relationshipId = `rId-${kind}`;
    const part = {
      ...headerFooter(kind),
      content: [
        {
          type: "paragraph" as const,
          content: [
            {
              type: "run" as const,
              content: [{ type: "text" as const, text: kind }],
            },
          ],
        },
      ],
    };
    const model: Document = {
      package: {
        document: { content: [] },
        ...(kind === "header"
          ? { headers: new Map([[relationshipId, part]]) }
          : { footers: new Map([[relationshipId, part]]) }),
      },
    };
    let readOnly = false;
    const manager = createHeaderFooterEditorManager({
      getDocument: () => model,
      getReadOnly: () => readOnly,
      getHost: () => host,
      getStyles: () => undefined,
      getTheme: () => undefined,
    });
    try {
      manager.sync();
      const view = manager.getView(relationshipId);
      if (!view)
        throw new Error("Expected a persistent header/footer editor view");

      readOnly = true;
      manager.sync();
      expect(view.editable).toBe(false);
      view.dispatch(view.state.tr.insertText(" forbidden", 1));
      expect(view.state.doc.textContent).toBe(kind);

      view.dispatch(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)),
      );
      expect(view.state.selection.anchor).toBe(3);

      readOnly = false;
      manager.sync();
      expect(manager.getView(relationshipId)).toBe(view);
      expect(view.editable).toBe(true);
      view.dispatch(
        view.state.tr.insertText(" editable", view.state.doc.content.size - 1),
      );
      expect(view.state.doc.textContent).toBe(`${kind} editable`);
    } finally {
      manager.destroy();
      host.remove();
    }
  },
);
