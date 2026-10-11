import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";

export const loadDirect = (page: Page, bytes: number[]) =>
  page.evaluate((input) => globalThis.__folioCanonical!.load(input), bytes);

export const loadOptional = (page: Page, bytes: number[]) =>
  page.evaluateHandle((input) => globalThis.__folioCanonical?.load(input), bytes);

export const loadComputed = (page: Page, bytes: number[]) =>
  page.evaluate((input) => globalThis["__folioCanonical"]?.["load"](input), bytes);
