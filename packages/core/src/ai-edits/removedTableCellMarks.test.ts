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

/**
 * A selection over a whole document holding a nested table, replaced while
 * suggesting (browser input seed 1794471204, path 14:0:4:3:3): every row and
 * paragraph mark is deleted, a nested table's rows included. The nested
 * table's paragraphs belong to its own rows and to the outer cell holding it;
 * accepting removes both, settling each paragraph's mark once.
 */
const NESTED_SAVED_BODY = [
  `<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="0" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Before the table.</w:delText></w:r></w:del></w:p>`,
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/></w:tblGrid>`,
  `<w:tr><w:trPr><w:del w:id="3" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="4" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="5" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Merged A1</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="6" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="7" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell C1</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="8" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="9" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell D1</w:delText></w:r></w:del></w:p></w:tc></w:tr>`,
  `<w:tr><w:trPr><w:del w:id="10" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>`,
  `<w:p></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="11" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="12" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell C2</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr>`,
  `<w:p></w:p></w:tc></w:tr>`,
  `<w:tr><w:trPr><w:del w:id="13" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:trPr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="14" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="15" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell A3</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="16" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="17" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell B3</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="18" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="19" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Outer</w:delText></w:r></w:del></w:p>`,
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1400"/><w:gridCol w:w="1400"/></w:tblGrid>`,
  `<w:tr><w:trPr><w:del w:id="20" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:trPr><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="21" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="22" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Inner 1</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="23" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="24" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Inner 2</w:delText></w:r></w:del></w:p></w:tc></w:tr></w:tbl>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="25" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="26" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="27" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell D3</w:delText></w:r></w:del></w:p></w:tc></w:tr>`,
  `<w:tr><w:trPr><w:del w:id="28" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:trPr><w:tc><w:tcPr><w:vMerge/></w:tcPr>`,
  `<w:p></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="29" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="30" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell B4</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="31" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="32" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell C4</w:delText></w:r></w:del></w:p></w:tc><w:tc>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="33" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="34" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>Cell D4</w:delText></w:r></w:del></w:p></w:tc></w:tr></w:tbl>`,
  `<w:p><w:del w:id="35" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:delText>After the table.</w:delText></w:r></w:del><w:ins w:id="2" w:author="Reviewer" w:date="2026-11-21T00:00:00Z"><w:r><w:t>alph</w:t></w:r></w:ins></w:p>`,
].join("");

const replacedDocument = async () =>
  (
    await buildRawPackage({
      styles: null,
      body: NESTED_SAVED_BODY + SECTION(),
    })
  ).slice().buffer;

const NESTED_OUTCOMES = {
  accept: ["alph"],
  reject: [
    "Before the table.",
    "Merged A1",
    "Cell C1",
    "Cell D1",
    "Cell C2",
    "Cell A3",
    "Cell B3",
    "Outer",
    "Inner 1",
    "Inner 2",
    "",
    "Cell D3",
    "Cell B4",
    "Cell C4",
    "Cell D4",
    "After the table.",
  ],
} as const;

for (const session of ["legacy", "canonical"] as const)
  for (const decision of ["accept", "reject"] as const)
    test(`${decision}ing a whole document replaced over a nested table resolves every change (${session})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await replacedDocument(),
        { session },
      );
      if (decision === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      expect(reviewer.getChanges()).toEqual([]);
      expect(reviewer.getContent().map(({ text }) => text)).toEqual([
        ...NESTED_OUTCOMES[decision],
      ]);
    });
