import { describe, expect, test } from "bun:test";

import type { Document } from "@stll/folio-core/types/document";

import { describePackageDifference } from "./lib/corpus-invariants/model-equality";

/**
 * The model-equality projection is data-shaped, so a test can hand it a package
 * literal. Nothing here parses a file: what is under test is which differences
 * survive normalisation, not how a package is read.
 */
const documentWith = (content: unknown): Document =>
  // SAFETY: the projection walks plain data and never reads a typed field; a
  // full parsed package would say nothing more about which keys it erases.
  ({ package: { document: { content } } }) as unknown as Document;

describe("describePackageDifference", () => {
  test("two identical packages differ in nothing", () => {
    expect(
      describePackageDifference(documentWith([{ text: "a" }]), documentWith([{ text: "a" }])),
    ).toBeNull();
  });

  test("a changed field is reported with its path and both values", () => {
    expect(
      describePackageDifference(documentWith([{ bold: true }]), documentWith([{ bold: false }])),
    ).toBe("package.document.content[].bold: true became false");
  });

  test("array positions collapse, so one defect is one signature", () => {
    const left = describePackageDifference(
      documentWith([{ bold: true }, { bold: true }]),
      documentWith([{ bold: true }, { bold: false }]),
    );
    const right = describePackageDifference(
      documentWith([{ bold: true }]),
      documentWith([{ bold: false }]),
    );
    expect(left).toBe(right);
  });

  /**
   * The slot a capture lives in is not content, and the invariant that forces
   * serialization removes it by construction. Erasing it to a sentinel rather
   * than dropping the key would make the removal itself the difference, on
   * every package that carries the slot.
   */
  test("a capture slot present on one side only is not a difference", () => {
    expect(
      describePackageDifference(
        documentWith([{ text: "a", sourceXml: "<w:tblPr/>" }]),
        documentWith([{ text: "a" }]),
      ),
    ).toBeNull();
    expect(
      describePackageDifference(
        documentWith([{ text: "a", rawEndPropertiesXml: "<w:sdtEndPr/>" }]),
        documentWith([{ text: "a" }]),
      ),
    ).toBeNull();
  });

  test("a volatile field present on one side only is not a difference", () => {
    expect(
      describePackageDifference(
        documentWith([{ text: "a", lastModifiedBy: "someone" }]),
        documentWith([{ text: "a" }]),
      ),
    ).toBeNull();
  });

  test("a capture slot with different content on both sides is still not a difference", () => {
    expect(
      describePackageDifference(
        documentWith([{ sourceXml: "<w:tblPr><w:x/></w:tblPr>" }]),
        documentWith([{ sourceXml: "<w:tblPr/>" }]),
      ),
    ).toBeNull();
  });

  test("a long string is reported by type, so no document text reaches a signature", () => {
    const difference = describePackageDifference(
      documentWith([{ text: "a".repeat(200) }]),
      documentWith([{ text: "b".repeat(200) }]),
    );
    expect(difference).toBe("package.document.content[].text: string became string");
  });

  test("a media buffer compares by length, not by bytes", () => {
    expect(
      describePackageDifference(
        documentWith([{ image: new Uint8Array([1, 2, 3]) }]),
        documentWith([{ image: new Uint8Array([9, 9, 9]) }]),
      ),
    ).toBeNull();
  });

  test("a map compares by sorted entries, not by insertion order", () => {
    expect(
      describePackageDifference(
        documentWith([
          {
            headers: new Map([
              ["rId1", "a"],
              ["rId2", "b"],
            ]),
          },
        ]),
        documentWith([
          {
            headers: new Map([
              ["rId2", "b"],
              ["rId1", "a"],
            ]),
          },
        ]),
      ),
    ).toBeNull();
  });
});
