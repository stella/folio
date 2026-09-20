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
 * A loss belongs to the element that held it, and a path that names only the
 * field it happened to reports two owners as one defect. A run and the
 * paragraph around it both carry `preservedAttributes`, and only the run's is
 * a decision the container contract made. While a segment named only the
 * field, the two read alike, so the disposition that claims the run's could
 * only approximate the paragraph's away by the shape of the path, and the
 * paragraph-owned rows that happened to fit that shape sat inside a count
 * nobody could tell apart.
 */
describe("a path segment names the kind it steps into", () => {
  const run = (fields: Record<string, unknown> = {}) => ({ type: "run", ...fields });
  const paragraph = (content: unknown[], fields: Record<string, unknown> = {}) => ({
    type: "paragraph",
    content,
    ...fields,
  });

  test("an element the model discriminates carries its kind", () => {
    expect(
      onlyMessage([paragraph([run({ preservedAttributes: [1] })])], [paragraph([run()])]),
    ).toBe(
      "package.document.content[paragraph].content[run].preservedAttributes: array became absent",
    );
  });

  test("a run-owned and a paragraph-owned loss of the same field are two signatures", () => {
    const runOwned = onlyMessage(
      [paragraph([run({ preservedAttributes: [1] })])],
      [paragraph([run()])],
    );
    const paragraphOwned = onlyMessage(
      [{ type: "blockSdt", content: [paragraph([], { preservedAttributes: [1] })] }],
      [{ type: "blockSdt", content: [paragraph([])] }],
    );
    expect(runOwned).toBe(
      "package.document.content[paragraph].content[run].preservedAttributes: array became absent",
    );
    expect(paragraphOwned).toBe(
      "package.document.content[blockSdt].content[paragraph].preservedAttributes: array became absent",
    );
    expect(runOwned).not.toBe(paragraphOwned);
  });

  test("an element the model gives no discriminator stays untyped", () => {
    // A section is a model member without a `type`, so there is no kind to name.
    expect(onlyMessage([{ properties: { a: 1 } }], [{ properties: { a: 2 } }])).toBe(
      "package.document.content[].properties.a: 1 became 2",
    );
  });

  test("the kind is read from whichever side declares one", () => {
    expect(onlyMessage([run({ bold: true })], [{ bold: false }])).toBe(
      "package.document.content[run].bold: true became false",
    );
    expect(onlyMessage([{ bold: true }], [run({ bold: false })])).toBe(
      "package.document.content[run].bold: true became false",
    );
  });

  test("a length row belongs to the array, not to any one element", () => {
    expect(messagesBetween([paragraph([run(), run()])], [paragraph([run()])])).toEqual([
      "package.document.content[paragraph].content[]: length changed",
    ]);
  });

  /**
   * `type` is an ordinary key, and a package folio did not write can carry any
   * string under it. The vocabulary is closed for that reason: a segment names
   * a kind the model declares or it names none, so no path can spell a value
   * the document supplied.
   */
  test.each([
    ["a sentence from the document", "Please check this clause before Friday"],
    ["a name", "Jane Q. Reviewer"],
    ["a plausible near-miss", "runs"],
  ])("%s under `type` never reaches the path", (_label, hostile) => {
    const message = onlyMessage([{ type: hostile, bold: true }], [{ type: hostile, bold: false }]);
    expect(message).toBe("package.document.content[].bold: true became false");
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
