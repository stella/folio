import { describe, expect, test } from "bun:test";

import type { Document } from "@stll/folio-core/types/document";

import {
  describeChange,
  describePackageDifferences,
  differenceFailures,
  MAX_REPORTED_DIFFERENCES,
  omittedDifferencesMessage,
} from "./lib/corpus-invariants/model-equality";
import { valueVocabulary } from "./lib/corpus-invariants/value-vocabulary";
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
/**
 * `corpus/README.md`: "No corpus content is ever committed." A signature is
 * committed data, so a value read out of a document may not reach one. The
 * tokens the format itself defines are a different matter: they are the defect
 * (`"start"` became `"left"`), and they come from a schema, not from a file.
 */
describe("signatures quote tokens, never document content", () => {
  /**
   * A shape token is spelled in letters, so "carries none of the input" cannot
   * mean "shares no character with it". Three consecutive characters is the
   * operational bound: no fragment of the value a document supplied survives
   * into the row.
   */
  const carriesNoFragmentOf = (message: string, input: string): boolean =>
    Array.from({ length: Math.max(input.length - 2, 0) }, (_, index) =>
      input.slice(index, index + 3),
    ).every((fragment) => !message.includes(fragment));

  test.each([
    ["an author name", "author", "Jane Q. Reviewer", "Ana Nov", '"<string A>" became "<string B>"'],
    [
      "comment text",
      "text",
      "Please check this clause",
      "I checked it",
      '"<string A>" became "<string B>"',
    ],
    ["a drawing name", "name", "Grupo 17", "Gruppe 4", '"<string A>" became "<string B>"'],
    ["a style id", "styleId", "Ttulo1", "berschrift1", '"<string A>" became "<string B>"'],
    [
      "a hyperlink target",
      "href",
      "https://example.invalid/a",
      "https://example.invalid/b",
      '"<url A>" became "<url B>"',
    ],
  ])("%s never reaches the message", (_label, key, left, right, expected) => {
    const message = onlyMessage([{ [key]: left }], [{ [key]: right }]) ?? "";
    expect(message).toBe(`package.document.content[].${key}: ${expected}`);
    expect(carriesNoFragmentOf(message, left)).toBe(true);
    expect(carriesNoFragmentOf(message, right)).toBe(true);
  });

  test("two unrelated strings read as two strings, not as one unchanged value", () => {
    expect(onlyMessage([{ author: "Jane" }], [{ author: "Ana" }])).toBe(
      'package.document.content[].author: "<string A>" became "<string B>"',
    );
  });

  test("a value drawn from a closed set stays readable", () => {
    expect(onlyMessage([{ alignment: "start" }], [{ alignment: "left" }])).toBe(
      'package.document.content[].alignment: "start" became "left"',
    );
    expect(onlyMessage([{ conformance: "strict" }], [{ conformance: "transitional" }])).toBe(
      'package.document.content[].conformance: "strict" became "transitional"',
    );
    expect(onlyMessage([{ direction: "rtl" }], [{ direction: "ltr" }])).toBe(
      'package.document.content[].direction: "rtl" became "ltr"',
    );
    expect(onlyMessage([{ restart: "restart" }], [{ restart: "continue" }])).toBe(
      'package.document.content[].restart: "restart" became "continue"',
    );
  });

  test("a part path, a relationship id and a paragraph id read by shape", () => {
    expect(onlyMessage([{ target: "media/image1.png" }], [{ target: undefined }])).toBe(
      'package.document.content[].target: "<path>" became absent',
    );
    expect(onlyMessage([{ rId: "rId7" }], [{ rId: undefined }])).toBe(
      'package.document.content[].rId: "<id>" became absent',
    );
    expect(onlyMessage([{ paraId: undefined }], [{ paraId: "1A2B3C4D" }])).toBe(
      'package.document.content[].paraId: absent became "<hex>"',
    );
  });

  /**
   * The formatter is pure and the vocabulary is derived, so one golden set
   * pins both: a token that leaves either source changes this list, and a
   * value that starts being quoted verbatim changes it too.
   */
  test("the formatter is a pure function over a representative difference set", () => {
    const pairs: ReadonlyArray<readonly [unknown, unknown]> = [
      [undefined, null],
      [1, 2],
      [true, false],
      ["start", "left"],
      ["rtl", "ltr"],
      ["strict", "transitional"],
      ["Jane Q. Reviewer", "Ana"],
      ["media/image1.png", "media/image2.png"],
      ["https://example.invalid/a", "mailto:someone@example.invalid"],
      ["rId7", "rId8"],
      ["1A2B3C4D", "5E6F7A8B"],
      ["{2E4A9F1B-0000-4000-8000-1234567890AB}", undefined],
      [{ a: 1 }, [1]],
      ["a".repeat(200), "b".repeat(200)],
    ];
    expect(pairs.map(([left, right]) => describeChange(left, right))).toEqual([
      "absent became null",
      "1 became 2",
      "true became false",
      '"start" became "left"',
      '"rtl" became "ltr"',
      '"strict" became "transitional"',
      '"<string A>" became "<string B>"',
      '"<path A>" became "<path B>"',
      '"<url A>" became "<url B>"',
      '"<id A>" became "<id B>"',
      '"<hex A>" became "<hex B>"',
      '"<guid>" became absent',
      "object became array",
      '"<string A>" became "<string B>"',
    ]);
    // Called twice, same answer: nothing here reads a clock, a file or a
    // counter, so a census merged from four shards agrees with one run.
    expect(pairs.map(([left, right]) => describeChange(left, right))).toEqual(
      pairs.map(([left, right]) => describeChange(left, right)),
    );
  });
});

describe("the quoting vocabulary is derived, not written down", () => {
  test("it carries the schema's enumerations and folio's own closed sets", () => {
    const vocabulary = valueVocabulary();
    // From `specifications/generated/docx-transitional-schema.gen.json`.
    expect(vocabulary.has("transitional")).toBe(true);
    expect(vocabulary.has("continue")).toBe(true);
    expect(vocabulary.has("ltr")).toBe(true);
    // From `@stll/docx-core/model`: `DOCX_CONFORMANCE_CLASSES` adds this one,
    // and no schema enumeration declares it.
    expect(vocabulary.has("unknown")).toBe(true);
    expect(vocabulary.has("preserveOnly")).toBe(true);
  });

  test("no member carries whitespace, so no member can be a sentence", () => {
    expect([...valueVocabulary()].filter((token) => /\s/u.test(token))).toEqual([]);
  });
});
