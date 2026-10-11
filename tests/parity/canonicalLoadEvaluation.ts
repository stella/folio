import type { Page, PlaywrightWorkerOptions } from "@playwright/test";
import { TaggedError } from "better-result";
import type {} from "./canonicalBridge";

const CANONICAL_LOAD_TRANSPORTS = {
  chromium: { type: "cdp" },
  firefox: { type: "unsupported" },
  webkit: { type: "unsupported" },
} as const satisfies Record<
  PlaywrightWorkerOptions["browserName"],
  { type: "cdp" | "unsupported" }
>;

export class CanonicalLoadTransportError extends TaggedError("CanonicalLoadTransportError")<{
  message: string;
  engine: string;
}> {}

const isCanonicalLoadEngine = (engine: string): engine is keyof typeof CANONICAL_LOAD_TRANSPORTS =>
  Object.hasOwn(CANONICAL_LOAD_TRANSPORTS, engine);

const getCanonicalLoadTransport = (engine: string) => {
  if (!isCanonicalLoadEngine(engine))
    throw new CanonicalLoadTransportError({
      engine,
      message: `Unknown canonical load browser engine: ${engine}`,
    });
  return CANONICAL_LOAD_TRANSPORTS[engine];
};

/** Retain the exact promise CDP awaits, including across Chromium collection. */
export const evaluateCanonicalLoad = async (page: Page, bytes: number[]) => {
  const engine = page.context().browser()?.browserType().name() ?? "unavailable";
  const transport = getCanonicalLoadTransport(engine);
  switch (transport.type) {
    case "unsupported":
      throw new CanonicalLoadTransportError({
        engine,
        message: `Canonical load transport requires Chromium; ${engine} is unsupported`,
      });
    case "cdp":
      break;
    default: {
      const exhaustive: never = transport;
      return exhaustive;
    }
  }

  const session = await page.context().newCDPSession(page);
  let promiseObjectId: string | undefined;
  try {
    // Retain the pending promise independently of incidental parser or adapter
    // roots, then release its remote handle after the awaited result arrives.
    const operation = await session.send("Runtime.evaluate", {
      expression: `globalThis.__folioCanonical?.load(${JSON.stringify(bytes)})`,
      awaitPromise: false,
    });
    promiseObjectId = operation.result.objectId;
    if (operation.exceptionDetails || operation.result.subtype !== "promise" || !promiseObjectId)
      throw new CanonicalLoadTransportError({
        engine,
        message: operation.exceptionDetails?.text ?? "Canonical load did not return a promise",
      });

    const completion = await session.send("Runtime.awaitPromise", {
      promiseObjectId,
      returnByValue: true,
    });
    if (completion.exceptionDetails)
      throw new CanonicalLoadTransportError({
        engine,
        message:
          completion.exceptionDetails.exception?.description ?? completion.exceptionDetails.text,
      });
    if (typeof completion.result.value !== "boolean")
      throw new CanonicalLoadTransportError({
        engine,
        message: "Canonical load did not return a boolean result",
      });
    return completion.result.value;
  } finally {
    try {
      if (promiseObjectId)
        await session.send("Runtime.releaseObject", { objectId: promiseObjectId });
    } finally {
      await session.detach();
    }
  }
};
