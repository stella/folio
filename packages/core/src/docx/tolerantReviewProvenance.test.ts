/**
 * Review provenance is advisory metadata in an ignorable namespace. A value
 * the reader cannot read (another version, an unknown field, a malformed
 * value) is dropped with a warning; the document opens and its revisions are
 * read, resolved and rejected as they are.
 */

import { describe, expect, test } from "bun:test";

import { escapeXmlAttribute } from "@stll/docx-core";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "./parser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { createEmptyDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const attribute = (value: string) => escapeXmlAttribute(value);
const envelope = (value: unknown) => attribute(JSON.stringify(value));
const JOINS = { before: 0, after: 0, remove: 0 };

/** Provenance values an older or newer folio, or a damaged file, may carry. */
const UNREADABLE = {
  "a later version": envelope({ version: 2, value: JOINS }),
  "an unknown envelope field": envelope({
    version: 1,
    value: JOINS,
    extra: true,
  }),
  "an unknown value field": envelope({
    version: 1,
    value: { ...JOINS, extra: 1 },
  }),
  "a value of the wrong shape": envelope({ version: 1, value: "joins" }),
  "text that is not JSON": attribute("{not json"),
} as const;

/** A paragraph "ab" with "X" inserted between, the insertion carrying `markup`. */
const packageWith = async (insertionAttributes: string, runChange = "") => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}" mc:Ignorable="folio"><w:body><w:p><w:r><w:t>a</w:t></w:r><w:ins w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"${insertionAttributes}><w:r>${runChange}<w:t>X</w:t></w:r></w:ins><w:r><w:t>b</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" });
};

const outcome = async (bytes: Uint8Array) => {
  const document = await parseDocx(bytes.slice().buffer, {
    preloadFonts: false,
  });
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return reviewer.getContent().map(({ text }) => text);
  };
  return {
    dropped: (document.parseWarnings ?? []).filter(
      ({ code }) => code === PARSE_WARNING_CODES.reviewProvenanceDropped,
    ).length,
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

describe("unreadable review provenance", () => {
  for (const [kind, value] of Object.entries(UNREADABLE)) {
    test(`resolution joins with ${kind} are dropped and the revision kept`, async () => {
      expect(
        await outcome(await packageWith(` folio:resolutionJoins="${value}"`)),
      ).toEqual({ dropped: 1, accepted: ["aXb"], rejected: ["ab"] });
    });

    test(`a split piece with ${kind} is dropped and the revision kept`, async () => {
      expect(
        await outcome(
          await packageWith(` folio:splitResolutionJoins="${value}"`),
        ),
      ).toEqual({ dropped: 1, accepted: ["aXb"], rejected: ["ab"] });
    });

    test(`run property boundary joins with ${kind} are dropped and the change kept`, async () => {
      const runChange = `<w:rPr><w:b/><w:rPrChange w:id="2" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" folio:boundaryJoins="${value}"><w:rPr/></w:rPrChange></w:rPr>`;
      expect(await outcome(await packageWith("", runChange))).toEqual({
        dropped: 1,
        accepted: ["aXb"],
        rejected: ["ab"],
      });
    });
  }

  test("readable provenance reports nothing", async () => {
    expect(
      await outcome(
        await packageWith(
          ` folio:resolutionJoins="${envelope({ version: 1, value: JOINS })}"`,
        ),
      ),
    ).toEqual({ dropped: 0, accepted: ["aXb"], rejected: ["ab"] });
  });
});
