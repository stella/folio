import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";

export const loadCapturedAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical!;
    const start = (bytes: number[]) => bridge.load(bytes);
    return start(input);
  }, bytes);

export const loadGlobalDespiteShadow = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const start = (__folioCanonical: number[]) =>
      globalThis.__folioCanonical!.load(__folioCanonical);
    return start(input);
  }, bytes);

export const loadUnshadowedDestructuredAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const { load } = globalThis.__folioCanonical!;
    const start = (bytes: number[]) => load(bytes);
    return start(input);
  }, bytes);

export const loadNamedCallback = (page: Page, bytes: number[]) => {
  const loadInBrowser = (input: number[]) => globalThis.__folioCanonical!.load(input);
  return page.evaluate(loadInBrowser, bytes);
};
