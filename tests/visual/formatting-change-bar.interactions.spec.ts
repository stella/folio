import { expect } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createEmptyDocx } from "../../packages/core/src/docx/rezip";
import { reactPort, vuePort } from "../parity/canonicalInteractionAssertions";
import { test } from "../parity/canonicalTest";

/**
 * A pending formatting change, of a run's properties or of a paragraph's,
 * paints the text as it now is and marks its paragraph with the change bar
 * in the margin, in All Markup and Simple Markup, in both hosts and both
 * sessions. A paragraph with no change has no bar.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const STAMP = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

const packageBytes = async () => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  const body = [
    `<w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="1" ${STAMP}><w:rPr/></w:rPrChange></w:rPr><w:t>bolded</w:t></w:r></w:p>`,
    `<w:p w14:paraId="10000002"><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="2" ${STAMP}><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>`,
    `<w:p w14:paraId="10000003"><w:r><w:t>Untouched</w:t></w:r></w:p>`,
  ].join("");
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return [...(await zip.generate({ compression: "store" }))];
};

const HOSTS = [
  { name: "React", port: reactPort },
  { name: "Vue", port: vuePort },
] as const;
const SESSIONS = [
  { name: "default", query: "" },
  { name: "canonical", query: "?session=canonical" },
] as const;

test("a pending formatting change marks its paragraph with the change bar", async ({
  page,
}) => {
  test.setTimeout(120_000);
  const bytes = await packageBytes();
  for (const host of HOSTS)
    for (const session of SESSIONS) {
      const label = `${host.name}, ${session.name}`;
      await page.goto(`http://localhost:${host.port}/${session.query}`);
      await page.waitForSelector(".layout-page");
      expect(
        await page.evaluate(
          (data) => globalThis.__folioCanonical?.loadAnySession(data),
          bytes,
        ),
        label,
      ).toBe(true);
      await expect(page.locator(".layout-page").first(), label).toContainText(
        "Untouched",
      );
      const bars = () =>
        page.evaluate(() =>
          ["10000001", "10000002", "10000003"].map((paraId) => {
            const fragment = document.querySelector<HTMLElement>(
              `.layout-paragraph[data-para-id="${paraId}"]`,
            );
            if (!fragment) return [paraId, "missing"];
            const bar = getComputedStyle(fragment, "::before");
            return [
              paraId,
              fragment.dataset["trackedChanges"] === "true" &&
              bar.content !== "none" &&
              bar.width !== "0px"
                ? "bar"
                : "none",
            ];
          }),
        );
      const expected = [
        ["10000001", "bar"],
        ["10000002", "bar"],
        ["10000003", "none"],
      ];
      expect(await bars(), `${label}, All Markup`).toEqual(expected);
      // Simple Markup shows the text as it would be accepted, and marks the
      // same paragraphs.
      if (host.name === "React") {
        await page
          .getByRole("combobox")
          .filter({ hasText: "All Markup" })
          .click();
        await page
          .getByRole("option", { name: "Simple Markup", exact: true })
          .click();
      } else {
        await page.locator(".review-controls__display").click();
        await page
          .locator(".review-controls__option", { hasText: "Simple Markup" })
          .click({ force: true });
      }
      await expect
        .poll(bars, { message: `${label}, Simple Markup` })
        .toEqual(expected);
    }
});
