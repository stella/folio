import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { applyDocumentOps, validateOpsDocument } from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import {
  createCanonicalHeaderFooterOperation,
  removeCanonicalHeaderFooterOperations,
  withCanonicalParagraphIds,
} from "./canonicalOperations";

setDefaultTimeout(propertyTestTimeout(5_000));

const documentWithStories = (): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000001", content: [] }],
      finalSectionProperties: {},
    },
    headers: new Map([
      [
        "rId_new_header_default",
        { type: "header", content: [{ type: "paragraph", paraId: "00000002", content: [] }] },
      ],
    ]),
    footnotes: [
      {
        type: "footnote",
        id: 1,
        content: [{ type: "paragraph", paraId: "00000003", content: [] }],
      },
    ],
  },
});

test("command allocation repairs collisions across stories and invalid identifiers without mutating input", () => {
  fc.assert(
    fc.property(
      fc.array(
        fc.constantFrom(
          undefined,
          "00000000",
          "00000001",
          "00000002",
          "00000003",
          "0000abcd",
          "0000ABCD",
          "FFFFFFFF",
        ),
        { minLength: 1, maxLength: 15 },
      ),
      (identities) => {
        const document = documentWithStories();
        const content = identities.map((paraId) =>
          paraId === undefined
            ? { type: "paragraph" as const, content: [] }
            : { type: "paragraph" as const, content: [], paraId },
        );
        const before = structuredClone(content);
        const allocated = withCanonicalParagraphIds(content, document);
        expect(content).toEqual(before);
        const operation = createCanonicalHeaderFooterOperation({
          document,
          position: "header",
          referenceType: "default",
        });
        expect(operation.story.rId).toBe("rId_new_header_default_2");
        const applied = applyDocumentOps(document, [{ ...operation, content: allocated }]);
        expect(applied.isOk()).toBe(true);
        if (applied.isOk()) expect(validateOpsDocument(applied.value.document).isOk()).toBe(true);
      },
    ),
    propertyConfig(),
  );
});

test("removal addresses every explicit binding of a shared header part", () => {
  const document = documentWithStories();
  document.package.document.finalSectionProperties = {
    headerReferences: [
      { type: "default", rId: "rId_new_header_default" },
      { type: "first", rId: "rId_new_header_default" },
    ],
  };
  const operations = removeCanonicalHeaderFooterOperations({
    document,
    position: "header",
    rId: "rId_new_header_default",
  });
  const applied = applyDocumentOps(document, operations);
  expect(applied.isOk()).toBe(true);
  if (applied.isOk()) {
    expect(applied.value.document.package.headers?.has("rId_new_header_default")).toBe(false);
    expect(
      applied.value.document.package.document.finalSectionProperties?.headerReferences,
    ).toEqual([]);
  }
});
