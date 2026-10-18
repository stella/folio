import { describe, expect, test } from "bun:test";
import { setDocumentBodyComments } from "@stll/docx-core/model";

import type { Document } from "../types/document";
import { validateFolioDocumentModel } from "./modelValidation";

const documentWith = (body: Document["package"]["document"]): Document => ({
  package: { document: body },
});

describe("canonical Folio document model validation", () => {
  test("rejects an empty comments array", () => {
    const result = validateFolioDocumentModel(documentWith({ content: [], comments: [] }));

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "package.document.comments",
      message: "Empty comment lists must be absent.",
      severity: "error",
    });
  });

  test("rejects empty paragraph formatting", () => {
    const result = validateFolioDocumentModel(
      documentWith({ content: [{ type: "paragraph", content: [], formatting: {} }] }),
    );

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "package.document.content[0].formatting",
      message: "Empty paragraph formatting must be absent.",
      severity: "error",
    });
  });

  test("the comments owner stores undefined for an empty list", () => {
    const body: Document["package"]["document"] = { content: [] };

    setDocumentBodyComments(body, []);

    expect(body.comments).toBeUndefined();
    expect(validateFolioDocumentModel(documentWith(body)).valid).toBe(true);
  });
});
