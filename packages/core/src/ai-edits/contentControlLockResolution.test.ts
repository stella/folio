/**
 * Accepting or rejecting a tracked change honours a content control's
 * `w:lock`, as editing does.
 *
 * A change whose resolution would change the content of a control locked with
 * `contentLocked` or `sdtContentLocked`, or remove one locked with `sdtLocked`
 * or `sdtContentLocked`, stays pending as it is: `acceptChange` and
 * `rejectChange` return `false`, and `acceptAll` and `rejectAll` resolve every
 * other change and report the held ones as `contentControlLock`.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import {
  acceptChange,
  changeResolutionLockViolation,
  findChangeAtPosition,
  rejectChange,
} from "../prosemirror/commands/comments";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const createDocx = async (bodyXml: string): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>${bodyXml}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const DATE = 'w:author="Author" w:date="2024-01-01T00:00:00Z"';
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inserted = (id: number, text: string) =>
  `<w:ins w:id="${id}" ${DATE}>${run(text)}</w:ins>`;
const deleted = (id: number, text: string) =>
  `<w:del w:id="${id}" ${DATE}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const control = (lock: string | null, content: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="Guarded"/>` +
  (lock === null ? "" : `<w:lock w:val="${lock}"/>`) +
  `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A paragraph with its own insertion (id 9) outside any control. */
const FREE_PARAGRAPH = `<w:p>${run("Free ")}${inserted(9, "addition")}</w:p>`;

type Fixture = {
  name: string;
  body: (lock: string | null) => string;
  /** The revision ids inside, or around, the control. */
  heldIds: readonly number[];
  /** The locks under which resolving them breaks the lock, by mode. */
  holds: (lock: string | null, mode: "accept" | "reject") => boolean;
};

const locksContent = (lock: string | null) =>
  lock === "contentLocked" || lock === "sdtContentLocked";
const locksControl = (lock: string | null) =>
  lock === "sdtLocked" || lock === "sdtContentLocked";

const FIXTURES: readonly Fixture[] = [
  {
    name: "revisions inside an inline control",
    body: (lock) =>
      `<w:p>${run("Before ")}${control(lock, run("Keep") + inserted(1, " added") + deleted(2, " gone"))}${run(" after.")}</w:p>` +
      FREE_PARAGRAPH,
    heldIds: [1, 2],
    holds: (lock) => locksContent(lock),
  },
  {
    name: "revisions inside a block control",
    body: (lock) =>
      `<w:p>${run("Opening.")}</w:p>` +
      control(
        lock,
        `<w:p>${run("Keep")}${inserted(1, " added")}${deleted(2, " gone")}</w:p>`,
      ) +
      FREE_PARAGRAPH,
    heldIds: [1, 2],
    holds: (lock) => locksContent(lock),
  },
  {
    name: "a deletion around a whole inline control",
    body: (lock) =>
      `<w:p>${run("Before ")}<w:del w:id="3" ${DATE}>${control(lock, "<w:r><w:delText>CTRL</w:delText></w:r>")}</w:del>${run(" after.")}</w:p>` +
      FREE_PARAGRAPH,
    heldIds: [3],
    // Accepting removes the control, which only a control lock forbids;
    // rejecting changes what its text is, which a content lock forbids.
    holds: (lock, mode) =>
      mode === "accept" ? locksControl(lock) : locksContent(lock),
  },
];

const LOCKS = [
  null,
  "unlocked",
  "sdtLocked",
  "contentLocked",
  "sdtContentLocked",
] as const;

const documentXml = async (reviewer: FolioDocxReviewer): Promise<string> =>
  (await (
    await ZipArchive.read(await reviewer.toBuffer())
  )
    .entry("word/document.xml")
    ?.text()) ?? "";

// A save writes minted paragraph ids into the paragraphs of a part it
// rewrites; they are no content a lock holds.
const controlXml = (xml: string): string | undefined =>
  xml
    .match(/<w:sdt>.*?<\/w:sdt>/u)?.[0]
    .replaceAll(/ w14:(?:paraId|textId)="[0-9A-F]{8}"/gu, "");

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

type OpenOptions = {
  fixture: Fixture;
  lock: string | null;
  session: HeadlessReviewerSession;
};

const open = async ({ fixture, lock, session }: OpenOptions) =>
  FolioDocxReviewer.fromBuffer(await createDocx(fixture.body(lock)), {
    session,
  });

describe("resolving tracked changes in a locked content control", () => {
  for (const session of SESSIONS)
    for (const fixture of FIXTURES)
      for (const lock of LOCKS)
        for (const mode of ["accept", "reject"] as const) {
          const held = fixture.holds(lock, mode);
          const label = `${session}: ${fixture.name}, ${String(lock)}: ${mode}`;

          test(`${label} all ${held ? "leaves them pending" : "resolves them"}`, async () => {
            const reviewer = await open({ fixture, lock, session });
            const before = controlXml(await documentXml(reviewer));
            const result =
              mode === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
            const remaining = reviewer.getChanges().map(({ id }) => id);
            const xml = await documentXml(reviewer);
            // The free paragraph's insertion resolves either way.
            expect(remaining).not.toContain(9);
            expect(xml).toContain(mode === "accept" ? "addition" : "Free");
            if (held) {
              expect(result.skipped.map(({ id }) => id).toSorted()).toEqual([
                ...fixture.heldIds,
              ]);
              expect(
                result.skipped.every(
                  ({ reason, message }) =>
                    reason === "contentControlLock" &&
                    message.includes(String(lock)),
                ),
              ).toBe(true);
              expect(remaining.toSorted()).toEqual([...fixture.heldIds]);
              expect(controlXml(xml)).toBe(before);
            } else {
              expect(result.skipped).toEqual([]);
              expect(remaining).toEqual([]);
            }
          });

          test(`${label} one by one ${held ? "returns false" : "resolves it"}`, async () => {
            const reviewer = await open({ fixture, lock, session });
            const before = await documentXml(reviewer);
            const [first] = fixture.heldIds;
            if (first === undefined) throw new Error("expected a held id");
            const resolved =
              mode === "accept"
                ? reviewer.acceptChange(first)
                : reviewer.rejectChange(first);
            expect(resolved).toBe(!held);
            if (held)
              expect(controlXml(await documentXml(reviewer))).toBe(
                controlXml(before),
              );
            // A change the lock does not reach still resolves.
            expect(
              mode === "accept"
                ? reviewer.acceptChange(9)
                : reviewer.rejectChange(9),
            ).toBe(true);
          });
        }
});

describe("the change at a caret inside a locked inline control", () => {
  test("is found whole, and the editor commands refuse it", async () => {
    const fixture = FIXTURES[0];
    if (fixture === undefined) throw new Error("expected the inline fixture");
    const document = await parseDocx(
      await createDocx(fixture.body("contentLocked")),
      { preloadFonts: false },
    );
    const state = EditorState.create({ doc: toProseDoc(document) });
    let caret = -1;
    state.doc.descendants((node, position) => {
      if (node.isText && node.text === " added") caret = position + 2;
    });
    expect(caret).toBeGreaterThan(0);
    const range = findChangeAtPosition(state, caret, caret);
    expect(range.to - range.from).toBe(" added".length);
    for (const mode of ["accept", "reject"] as const) {
      expect(changeResolutionLockViolation(state, { ...range, mode })).toEqual({
        lock: "contentLocked",
        change: "editContent",
      });
      const command = (mode === "accept" ? acceptChange : rejectChange)(
        range.from,
        range.to,
      );
      let dispatched = false;
      expect(
        command(state, () => {
          dispatched = true;
        }),
      ).toBe(false);
      expect(dispatched).toBe(false);
    }
  });
});

describe("the mark of a paragraph holding a held change", () => {
  // Removing the mark runs the paragraph, with the locked control and the
  // held change in it, into the next one; the control's content is untouched.
  const removedMark = (mode: "accept" | "reject") =>
    `<w:pPr><w:rPr>${mode === "accept" ? `<w:del w:id="5" ${DATE}/>` : `<w:ins w:id="5" ${DATE}/>`}</w:rPr></w:pPr>`;
  for (const mode of ["accept", "reject"] as const)
    test(`${mode} all removes it and leaves only the held change`, async () => {
      const held =
        mode === "accept" ? inserted(1, " text") : deleted(1, " text");
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createDocx(
          `<w:p>${removedMark(mode)}${run("Held ")}${control("contentLocked", run("clause") + held)}</w:p>` +
            `<w:p>${run(" next.")}</w:p>`,
        ),
      );
      // `xml:space="preserve"` changes nothing on text with no space at its ends.
      const content = (xml: string) =>
        controlXml(xml)?.replaceAll(
          /<w:t xml:space="preserve">(\S[^<]*\S)<\/w:t>/gu,
          "<w:t>$1</w:t>",
        );
      const before = content(await documentXml(reviewer));
      const result =
        mode === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
      expect(result.skipped.map(({ id }) => id)).toEqual([1]);
      expect(reviewer.getChanges().map(({ id }) => id)).toEqual([1]);
      const xml = await documentXml(reviewer);
      expect(xml.match(/<w:p[ >]/gu)).toHaveLength(1);
      expect(xml).toContain(" next.");
      expect(content(xml)).toBe(before);
    });
});
