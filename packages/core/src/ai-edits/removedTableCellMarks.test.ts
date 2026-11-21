/**
 * A table pasted at the end of the document while suggesting, as the default
 * editor saves it (browser input seed 1794637615, path 0; its body as saved,
 * the comment and pasted symbols left out): the pasted text replaces the
 * document's words, the table is an inserted row whose cells' paragraph
 * marks are tracked too, and its last cell is an inserted cell ending in a
 * paragraph whose mark is inserted. Rejecting every change removes the table
 * with its cells' pending marks.
 */
import { expect, test } from "bun:test";

import { buildRawPackage, SECTION } from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "./headless";

const SAVED_BODY = [
  `<w:p><w:pPr><w:rPr><w:del w:id="2" w:author="Reviewer" w:date="2026-11-20T17:57:22.561Z"/></w:rPr></w:pPr><w:del w:id="1" w:author="Reviewer" w:date="2026-11-20T17:57:22.561Z"><w:r><w:delText>Intro paragraph.</w:delText></w:r></w:del></w:p>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="4" w:author="Reviewer" w:date="2026-11-20T17:57:22.561Z"/></w:rPr></w:pPr><w:del w:id="5" w:author="Reviewer" w:date="2026-11-20T17:57:22.561Z"><w:r><w:delText>Commented words</w:delText></w:r><w:r><w:delText xml:space="preserve"> sit in this paragraph.</w:delText></w:r></w:del></w:p>`,
  `<w:p><w:pPr><w:rPr><w:ins w:id="6" w:author="Reviewer" w:date="2026-11-20T17:57:22.816Z"/></w:rPr><w:pPrChange w:id="3" w:author="Reviewer" w:date="2026-11-20T17:57:22.562Z"><w:pPr/></w:pPrChange></w:pPr><w:del w:id="7" w:author="Reviewer" w:date="2026-11-20T17:57:22.561Z"><w:r><w:delText>Tail.</w:delText></w:r></w:del><w:ins w:id="8" w:author="Reviewer" w:date="2026-11-20T17:57:22.562Z"><w:r><w:t>shall</w:t></w:r></w:ins></w:p>`,
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>`,
  `<w:tr><w:trPr><w:ins w:id="9" w:author="Reviewer" w:date="2026-11-20T17:57:22.816Z"/></w:trPr>`,
  `<w:tc>`,
  `<w:p><w:pPr><w:rPr><w:ins w:id="10" w:author="Reviewer" w:date="2026-11-20T17:57:22.816Z"/></w:rPr></w:pPr><w:ins w:id="13" w:author="Reviewer" w:date="2026-11-20T17:57:22.816Z"><w:r><w:t>Left</w:t></w:r></w:ins></w:p></w:tc>`,
  `<w:tc>`,
  `<w:p><w:pPr><w:rPr><w:ins w:id="11" w:author="Reviewer" w:date="2026-11-20T17:57:22.982Z"/></w:rPr></w:pPr><w:ins w:id="14" w:author="Reviewer" w:date="2026-11-20T17:57:22.982Z"><w:r><w:t>Left</w:t></w:r></w:ins></w:p></w:tc>`,
  `<w:tc><w:tcPr><w:cellIns w:id="15" w:author="Reviewer" w:date="2026-11-20T17:57:22.982Z"/></w:tcPr>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="12" w:author="Reviewer" w:date="2026-11-20T17:57:22.983Z"/></w:rPr></w:pPr></w:p>`,
  `<w:p><w:pPr><w:rPr><w:ins w:id="16" w:author="Reviewer" w:date="2026-11-20T17:57:22.982Z"/></w:rPr></w:pPr><w:ins w:id="17" w:author="Reviewer" w:date="2026-11-20T17:57:22.982Z"><w:r><w:t>Right</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>`,
  `<w:p><w:pPr><w:pPrChange w:id="18" w:author="Reviewer" w:date="2026-11-20T17:57:22.562Z"><w:pPr/></w:pPrChange></w:pPr></w:p>`,
].join("");

const pastedTable = async () =>
  (
    await buildRawPackage({ styles: null, body: SAVED_BODY + SECTION() })
  ).slice().buffer;

const OUTCOMES = {
  reject: [
    "Intro paragraph.",
    "Commented words sit in this paragraph.",
    "Tail.",
  ],
  accept: ["shall", "Left", "Left", "Right", ""],
} as const;

for (const session of ["legacy", "canonical"] as const)
  for (const decision of ["accept", "reject"] as const)
    test(`${decision}ing a table pasted while suggesting resolves every change (${session})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await pastedTable(), {
        session,
      });
      if (decision === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      expect(reviewer.getChanges()).toEqual([]);
      expect(reviewer.getContent().map(({ text }) => text)).toEqual([
        ...OUTCOMES[decision],
      ]);
    });
