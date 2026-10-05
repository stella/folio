import { expect, test } from "@playwright/test";
import fc from "fast-check";
import { writeFileSync } from "node:fs";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { canonicalBrowserTraceArbitrary } from "./canonicalBrowserTrace";
import { driveBrowserIme, BrowserImeDriverCaptureError } from "./browserImeDriver";
import { Result } from "better-result";
import {
  installCanonicalHistoryProbe,
  beginCanonicalHistoryObservation,
  collectCanonicalHistoryObservation,
} from "./canonicalBrowserHistoryProbe";

test("seed 197 IME preamble records native completion and the next load", async ({
  page,
}, info) => {
  test.setTimeout(120_000);
  const action = fc
    .sample(canonicalBrowserTraceArbitrary, { seed: 197, path: "1", numRuns: 1 })
    .at(0)
    ?.at(0);
  if (action?.kind !== "imeReplacement" || action.completion !== "commit")
    throw new TypeError("Seed 197 preamble no longer begins with the committed IME gesture");
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  await page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors = [];
  });
  await installCanonicalHistoryProbe(page);
  const source = [
    ...new Uint8Array(await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }))),
  ];
  type CaseState = Awaited<
    ReturnType<NonNullable<typeof globalThis.__folioCanonical>["caseState"]>
  >;
  let before: CaseState | undefined;
  let after: CaseState | undefined;
  let reloaded: CaseState | undefined;
  const captures = new Map<
    string,
    Awaited<ReturnType<typeof collectCanonicalHistoryObservation>>
  >();
  let phase = "load";
  let failure: { message: string; primary?: string; observer?: string } | undefined;
  const finalCaptures: Awaited<ReturnType<typeof collectCanonicalHistoryObservation>>[] = [];
  try {
    expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source)).toBe(
      true,
    );
    expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
    before = await page.evaluate(() => globalThis.__folioCanonical?.caseState());
    expect(before?.composing).toBe(false);
    phase = "ime";
    await beginCanonicalHistoryObservation(page, {
      type: "input",
      index: 0,
      action: "imeReplacement",
    });
    await driveBrowserIme(page, action);
    // Observe delayed native completion too; this is measurement, not cleanup.
    await page.waitForTimeout(100);
    after = await page.evaluate(() => globalThis.__folioCanonical?.caseState());
    const ime = await collectCanonicalHistoryObservation(page);
    captures.set(phase, ime);
    expect(ime.capture.status).toBe("complete");
    if (ime.capture.status === "complete") {
      expect(ime.capture.driverCalls.map(({ operation }) => operation)).toEqual([
        "start",
        ...action.updates.slice(1).map(() => "update"),
        "commit",
      ]);
      expect(
        ime.capture.nativeEvents.some(
          (event) => event.type === "compositionstart" && event.target === "editor",
        ),
      ).toBe(true);
    }
    phase = "reload";
    await beginCanonicalHistoryObservation(page, { type: "load" });
    expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source)).toBe(
      true,
    );
    expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
    reloaded = await page.evaluate(() => globalThis.__folioCanonical?.caseState());
    captures.set(phase, await collectCanonicalHistoryObservation(page));
  } catch (cause) {
    failure = {
      message: cause instanceof Error ? cause.message : String(cause),
      ...(cause instanceof BrowserImeDriverCaptureError
        ? {
            primary: String(cause.cause),
            observer: String(cause.observationError),
          }
        : {}),
    };
    throw cause;
  } finally {
    const finalCapture = await Result.tryPromise({
      try: () => collectCanonicalHistoryObservation(page),
      catch: (cause: unknown) => cause,
    });
    if (finalCapture.isOk()) finalCaptures.push(finalCapture.value);
    const evidence = {
      seed: 197,
      path: "1",
      action,
      phase,
      before,
      after,
      reloaded,
      captures: Object.fromEntries(captures),
      finalCaptures,
      failure,
      finalCaptureError: finalCapture.isErr() ? String(finalCapture.error) : undefined,
    };
    const output = info.outputPath("canonical-ime-lifecycle.json");
    writeFileSync(output, JSON.stringify(evidence));
    await info.attach("canonical-ime-lifecycle", { path: output, contentType: "application/json" });
    console.log(`FOLIO_IME_LIFECYCLE ${JSON.stringify(evidence)}`);
  }
});
