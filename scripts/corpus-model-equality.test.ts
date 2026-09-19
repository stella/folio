import { describe, expect, test } from "bun:test";

import type { Document } from "@stll/folio-core/types/document";

import {
  describePackageDifferences,
  differenceFailures,
  MAX_REPORTED_DIFFERENCES,
  omittedDifferencesMessage,
} from "./lib/corpus-invariants/model-equality";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";

/**
 * The model-equality projection is data-shaped, so a test can hand it a package
 * literal. Nothing here parses a file: what is under test is which differences
 * survive normalisation, not how a package is read.
 */
const documentWith = (content: unknown): Document =>
  // SAFETY: the projection walks plain data and never reads a typed field; a
  // full parsed package would say nothing more about which keys it erases.
  ({ package: { document: { content } } }) as unknown as Document;

const messagesBetween = (left: unknown, right: unknown): readonly string[] =>
  describePackageDifferences(documentWith(left), documentWith(right)).messages;

const onlyMessage = (left: unknown, right: unknown): string | undefined =>
  messagesBetween(left, right).at(0);

describe("describePackageDifferences", () => {
  test("two identical packages differ in nothing", () => {
    expect(messagesBetween([{ text: "a" }], [{ text: "a" }])).toEqual([]);
  });

  test("a changed field is reported with its path and both values", () => {
    expect(onlyMessage([{ bold: true }], [{ bold: false }])).toBe(
      "package.document.content[].bold: true became false",
    );
  });

  test("array positions collapse, so one defect is one signature", () => {
    expect(onlyMessage([{ bold: true }, { bold: true }], [{ bold: true }, { bold: false }])).toBe(
      onlyMessage([{ bold: true }], [{ bold: false }]),
    );
  });

  /**
   * The slot a capture lives in is not content, and the invariant that forces
   * serialization removes it by construction. Erasing it to a sentinel rather
   * than dropping the key would make the removal itself the difference, on
   * every package that carries the slot.
   */
  test("a capture slot present on one side only is not a difference", () => {
    expect(messagesBetween([{ text: "a", sourceXml: "<w:tblPr/>" }], [{ text: "a" }])).toEqual([]);
    expect(
      messagesBetween([{ text: "a", rawEndPropertiesXml: "<w:sdtEndPr/>" }], [{ text: "a" }]),
    ).toEqual([]);
  });

  test("a volatile field present on one side only is not a difference", () => {
    expect(messagesBetween([{ text: "a", lastModifiedBy: "someone" }], [{ text: "a" }])).toEqual(
      [],
    );
  });

  test("a capture slot with different content on both sides is still not a difference", () => {
    expect(
      messagesBetween([{ sourceXml: "<w:tblPr><w:x/></w:tblPr>" }], [{ sourceXml: "<w:tblPr/>" }]),
    ).toEqual([]);
  });

  test("a long string is reported by type, so no document text reaches a signature", () => {
    expect(onlyMessage([{ text: "a".repeat(200) }], [{ text: "b".repeat(200) }])).toBe(
      "package.document.content[].text: string became string",
    );
  });

  test("a media buffer compares by length, not by bytes", () => {
    expect(
      messagesBetween(
        [{ image: new Uint8Array([1, 2, 3]) }],
        [{ image: new Uint8Array([9, 9, 9]) }],
      ),
    ).toEqual([]);
  });

  test("a map compares by sorted entries, not by insertion order", () => {
    expect(
      messagesBetween(
        [
          {
            headers: new Map([
              ["rId1", "a"],
              ["rId2", "b"],
            ]),
          },
        ],
        [
          {
            headers: new Map([
              ["rId2", "b"],
              ["rId1", "a"],
            ]),
          },
        ],
      ),
    ).toEqual([]);
  });
});

/**
 * The ratchet reads a signature that appears for the first time as a new
 * defect. While a file reported only its first difference, fixing that
 * difference revealed the next one and the gate failed the fix. These are the
 * three claims that stop being true.
 */
describe("every difference a file exhibits", () => {
  test("two independent differences in one file produce two messages", () => {
    expect(
      messagesBetween([{ bold: true, italic: true }], [{ bold: false, italic: false }]),
    ).toEqual([
      "package.document.content[].bold: true became false",
      "package.document.content[].italic: true became false",
    ]);
  });

  test("fixing one difference leaves the other reported exactly as before", () => {
    const before = messagesBetween(
      [{ bold: true, italic: true }],
      [{ bold: false, italic: false }],
    );
    const after = messagesBetween([{ bold: true, italic: true }], [{ bold: true, italic: false }]);
    expect(after).toEqual(["package.document.content[].italic: true became false"]);
    expect(before).toContain(after[0] as string);
    expect(before.length).toBe(after.length + 1);
  });

  test("the same difference under two blocks is reported once", () => {
    expect(
      messagesBetween([{ bold: true }, { bold: true }], [{ bold: false }, { bold: false }]),
    ).toEqual(["package.document.content[].bold: true became false"]);
  });

  /**
   * Comparing arrays of different lengths index by index reports the shift, not
   * the loss, so the length row stays the end of that subtree.
   */
  test("an array whose length changed is one message and is not descended into", () => {
    expect(messagesBetween([{ runs: [{ bold: true }] }], [{ runs: [] }])).toEqual([
      "package.document.content[].runs[]: length changed",
    ]);
  });

  test("past the cap the file reports a marker instead, and says how many it dropped", () => {
    const LETTERS = "abcdefghijklmnopqrstuvwxyz";
    // Alphabetic, because a signature erases every digit: `f1` and `f2` are one
    // row, and a cap counts rows.
    const fields = (value: boolean): Record<string, boolean> =>
      Object.fromEntries(
        Array.from({ length: MAX_REPORTED_DIFFERENCES + 10 }, (_, index) => [
          `${LETTERS[Math.floor(index / LETTERS.length)] ?? "z"}${LETTERS[index % LETTERS.length] ?? "z"}`,
          value,
        ]),
      );
    const { messages, omitted } = describePackageDifferences(
      documentWith([fields(true)]),
      documentWith([fields(false)]),
    );
    expect(messages.length).toBe(MAX_REPORTED_DIFFERENCES);
    expect(omitted).toBe(10);

    const failures = differenceFailures(
      EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
      { messages, omitted },
      (message) => message,
    );
    expect(failures.length).toBe(MAX_REPORTED_DIFFERENCES + 1);
    expect(failures.at(-1)?.message).toBe(
      // The marker is a signature like any other, so its count is erased the
      // way every count in a message is.
      omittedDifferencesMessage(10).replace("10", "N"),
    );
  });

  test("a comparison with nothing past the cap adds no marker", () => {
    const failures = differenceFailures(
      EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
      describePackageDifferences(documentWith([{ bold: true }]), documentWith([{ bold: false }])),
      (message) => message,
    );
    expect(failures.length).toBe(1);
  });
});
