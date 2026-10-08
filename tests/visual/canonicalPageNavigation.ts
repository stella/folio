import type { Page, Request } from "@playwright/test";
import type {} from "../parity/canonicalBridge";

type NavigationState =
  | { status: "settled" }
  | { status: "loading"; request: Request; load: Promise<unknown> | null }
  | { status: "failed"; reason: string };

const owners = new WeakMap<Page, { current: NavigationState }>();

/** Observe requests before a navigation commits and destroys the old context. */
export const observeCanonicalPageNavigation = (page: Page) => {
  const existing = owners.get(page);
  if (existing) return existing;
  const owner: { current: NavigationState } = { current: { status: "settled" } };
  owners.set(page, owner);
  page.on("request", (request) => {
    if (request.isNavigationRequest() && request.frame() === page.mainFrame())
      owner.current = { status: "loading", request, load: null };
  });
  page.on("requestfailed", (request) => {
    if (owner.current.status === "loading" && owner.current.request === request)
      owner.current = { status: "failed", reason: request.failure()?.errorText ?? "unknown" };
  });
  page.on("load", () => {
    owner.current = { status: "settled" };
  });
  return owner;
};

/** Wait for both the new document and the playground's mounted bridge. */
export const waitForCanonicalPageReady = async (page: Page) => {
  const owner = observeCanonicalPageNavigation(page);
  for (;;) {
    if (owner.current.status === "failed")
      throw new TypeError(`Canonical playground navigation failed: ${owner.current.reason}`);
    if (owner.current.status === "loading") {
      // Concurrent oracle reads share the same navigation barrier.
      owner.current.load ??= page.waitForEvent("load");
      await owner.current.load;
    }
    await page.waitForLoadState("load");
    await page.waitForFunction(
      () => document.readyState === "complete" && globalThis.__folioCanonicalReady === true,
    );
    // A request may start while readiness is being awaited. Observe its load
    // before invoking the callback; never retry an interrupted evaluation.
    if (owner.current.status === "settled") return;
  }
};

/** The oracle's only entry into an execution context, after navigation settles. */
export const evaluateCanonicalPage = async <T>(page: Page, evaluate: () => Promise<T>) => {
  await waitForCanonicalPageReady(page);
  return evaluate();
};
