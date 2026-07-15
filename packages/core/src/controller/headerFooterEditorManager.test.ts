import { describe, expect, test } from "bun:test";

import type { Document, HeaderFooter } from "../types/document";
import {
  enumerateDocumentHeaderFooterParts,
  enumerateHeaderFooterParts,
} from "./headerFooterEditorManager";

const headerFooter = (type: HeaderFooter["type"]): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [],
});

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
            { content: [], properties: { headerReferences: [{ type: "default", rId: "rIdH" }] } },
            { content: [], properties: { headerReferences: [{ type: "default", rId: "rIdH" }] } },
          ],
        },
        headers: new Map([["rIdH", headerFooter("header")]]),
      },
    };

    expect(enumerateDocumentHeaderFooterParts(document)).toEqual([{ kind: "header", rId: "rIdH" }]);
  });
});
