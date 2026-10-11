import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";

export const loadBridgeAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical;
    return bridge?.load(input);
  }, bytes);

export const loadMethodAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical!;
    const load = bridge.load;
    return load(input);
  }, bytes);

export const loadDestructuredAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const { load: startLoad } = globalThis.__folioCanonical!;
    return startLoad(input);
  }, bytes);

export const loadNestedCallback = (page: Page, bytes: number[]) =>
  page.evaluate((input) => ({ operation: globalThis.__folioCanonical?.load(input) }), bytes);

export const loadNestedFunction = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical!;
    const start = () => bridge.load(input);
    return start();
  }, bytes);

export const loadBoundAlias = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical!;
    const load = bridge.load.bind(bridge);
    return load(input);
  }, bytes);

export const loadWithCall = (page: Page, bytes: number[]) =>
  page.evaluate((input) => {
    const bridge = globalThis.__folioCanonical!;
    return bridge.load.call(bridge, input);
  }, bytes);
