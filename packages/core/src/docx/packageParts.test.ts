import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { reconcilePackageReferences, resolvePackageRelationshipTarget } from "./packageParts";

test(
  "encoded relationship targets resolve once and reconciliation retains their parts",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("my document.xml", "článek_日本.xml", "100%20literal.xml"),
        fc.constantFrom("", "/", "./", "../word/"),
        async (name, prefix) => {
          const target = `${prefix}${prefix === "/" ? "word/" : ""}${encodeURIComponent(name)}`;
          const path = `word/${name}`;
          expect(resolvePackageRelationshipTarget(target, "word/_rels/document.xml.rels")).toBe(
            path.toLowerCase(),
          );
          const zip = new JSZip();
          zip.file(path, "<part/>");
          zip.file(
            "word/_rels/document.xml.rels",
            `<Relationships><Relationship Id="r1" Target="${target}"/></Relationships>`,
          );
          expect((await reconcilePackageReferences(zip, 6)).danglingRelationships).toEqual([]);
        },
      ),
      { numRuns: 24 },
    );
  },
  propertyTestTimeout(30_000),
);

test.each([
  "../word/document.xml",
  "%2E%2E/word/document.xml",
  "word%2Fdocument.xml",
  "word/%5Cdocument.xml",
  "word/%00.xml",
  "word/%ZZ.xml",
  "word/%C3.xml",
  "//host/document.xml",
])("rejects invalid or escaping package URI %s", (target) =>
  expect(resolvePackageRelationshipTarget(target, "_rels/.rels")).toBeUndefined(),
);
