/**
 * Package relationships are written in one namespace, in Strict and
 * Transitional packages alike. A package `_rels/.rels` in another namespace
 * is malformed: both engines refuse it with one typed error naming the part.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { MalformedPackageRelationshipsError } from "../docx/packageRelationshipsRoot";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = ["canonical"] as const;
const PACKAGE_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_OFFICE_NAMESPACE =
  "http://purl.oclc.org/ooxml/officeDocument/relationships";

const withPackageRelationships = async (edit: (xml: string) => string) => {
  const zip = await ZipArchive.read(
    await createDocx(fromMarkdown("Plain words.")),
  );
  const xml = (await zip.entry("_rels/.rels")?.text()) ?? "";
  const edited = edit(xml);
  expect(edited).not.toBe(xml);
  zip.set("_rels/.rels", edited);
  return (await zip.generate()).slice().buffer;
};

const CASES = {
  "the office relationships namespace": {
    edit: (xml: string) =>
      xml.replace(
        `xmlns="${PACKAGE_NAMESPACE}"`,
        `xmlns="${OFFICE_NAMESPACE}"`,
      ),
    namespaceUri: OFFICE_NAMESPACE,
  },
  "the Strict office relationships namespace": {
    edit: (xml: string) =>
      xml.replace(
        `xmlns="${PACKAGE_NAMESPACE}"`,
        `xmlns="${STRICT_OFFICE_NAMESPACE}"`,
      ),
    namespaceUri: STRICT_OFFICE_NAMESPACE,
  },
  "no namespace": {
    edit: (xml: string) => xml.replace(` xmlns="${PACKAGE_NAMESPACE}"`, ""),
    namespaceUri: undefined,
  },
} as const;

describe("package relationships in a namespace OPC does not define", () => {
  for (const session of SESSIONS)
    for (const [name, { edit, namespaceUri }] of Object.entries(CASES))
      test(`${session}: ${name} refuses the package`, async () => {
        const opened = FolioDocxReviewer.fromBuffer(
          await withPackageRelationships(edit),
          { session },
        );
        await expect(opened).rejects.toBeInstanceOf(
          MalformedPackageRelationshipsError,
        );
        await expect(opened).rejects.toMatchObject({
          part: "_rels/.rels",
          namespaceUri,
        });
      });

  for (const session of SESSIONS)
    test(`${session}: a prefixed package namespace still opens`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await withPackageRelationships((xml) =>
          xml
            .replace(
              `xmlns="${PACKAGE_NAMESPACE}"`,
              `xmlns:pr="${PACKAGE_NAMESPACE}"`,
            )
            .replaceAll("<Relationship", "<pr:Relationship")
            .replaceAll("</Relationships", "</pr:Relationships"),
        ),
        { session },
      );
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
        "Plain words.",
      ]);
    });
});
