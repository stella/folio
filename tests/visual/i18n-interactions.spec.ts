/**
 * Localization (behaviour) regression test.
 *
 * Proves folio's bundled translations render end-to-end: the playground merges
 * `getFolioMessages(locale)` into an `IntlProvider` and the toolbar's Bold
 * button exposes its localized `aria-label`. Switching the language switcher
 * re-renders the chrome in the chosen locale; Arabic also flips the shell to
 * RTL. Asserts content/attributes only (no screenshots), so it is stable across
 * machines and runs in CI's `interactions` project.
 */

import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

// Stable toolbar Bold aria-label per locale (FormattingBar `ariaLabel={t("bold")}`).
const BOLD_LABEL = {
  en: "Bold",
  de: "Fett",
  ar: "عريض",
  "pt-BR": "Negrito",
  "zh-CN": "加粗",
} as const;

async function openPlayground(page: Page): Promise<void> {
  await page.goto("/");
  await page.waitForSelector('[data-testid="folio-editor"]', { timeout: 15_000 });
  // Default locale renders the English label.
  await expect(page.locator(`[aria-label="${BOLD_LABEL.en}"]`)).toBeVisible();
}

async function switchLocale(page: Page, locale: string): Promise<void> {
  await page.getByTestId("language-select").selectOption(locale);
}

test("toolbar localizes to German via the bundled catalog", async ({ page }) => {
  await openPlayground(page);

  await switchLocale(page, "de");

  await expect(page.locator(`[aria-label="${BOLD_LABEL.de}"]`)).toBeVisible();
  // The English label is gone: it rendered the translation, not the source key.
  await expect(page.locator(`[aria-label="${BOLD_LABEL.en}"]`)).toHaveCount(0);
});

for (const locale of ["pt-BR", "zh-CN"] as const) {
  test(`toolbar localizes to ${locale} via the bundled catalog`, async ({ page }) => {
    await openPlayground(page);

    await switchLocale(page, locale);

    await expect(page.locator(`[aria-label="${BOLD_LABEL[locale]}"]`)).toBeVisible();
    await expect(page.locator(`[aria-label="${BOLD_LABEL.en}"]`)).toHaveCount(0);
  });
}

test("toolbar localizes to Arabic and flips the shell to RTL", async ({ page }) => {
  await openPlayground(page);

  await switchLocale(page, "ar");

  await expect(page.locator(`[aria-label="${BOLD_LABEL.ar}"]`)).toBeVisible();
  await expect(page.locator("div.pg-shell")).toHaveAttribute("dir", "rtl");
});
