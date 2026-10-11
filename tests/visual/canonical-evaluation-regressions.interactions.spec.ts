import { expect } from "@playwright/test";
import { Result } from "better-result";
import { readFileSync } from "node:fs";
import { test } from "./canonicalTimerProbe";
import { waitForCanonicalPageReady } from "./canonicalPageNavigation";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import {
  MODIFIER,
  reactPort,
  vuePort,
  snapshot,
  clearRefusals,
  expectNoRefusals,
  select,
  loadReady,
} from "../parity/canonicalInteractionAssertions";
import {
  checkCanonicalBrowserHistory,
  initializeCanonicalBrowserHistory,
} from "./canonicalBrowserHistoryOracle";
import { CANONICAL_EVALUATION_HISTORY_REPLAY } from "./canonicalBrowserTrace";
import { canonicalHistoryReplayActions, runCanonicalHistoryReplay } from "./canonicalHistoryReplay";

test("canonical structural gestures and toolbar operations preserve each intermediate history state", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const source = await canonicalLoadFixture(
    await createDocx(createEmptyDocument({ initialText: "ab" })),
  );
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    const reload = async () => {
      await loadReady(page, source);
      await select(page, 2);
    };
    const undoExactly = async (before: Awaited<ReturnType<typeof snapshot>>) => {
      const after = await snapshot(page);
      expect(after?.document).not.toEqual(before?.document);
      expect(after?.projectionMatchesCanonical).toBe(true);
      await page.keyboard.press(`${MODIFIER}+z`);
      const undone = await snapshot(page);
      expect(undone?.document).toEqual(before?.document);
      expect(undone?.selection).toEqual(before?.selection);
      expect(undone?.projectionJSON).toEqual(before?.projectionJSON);
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      const redone = await snapshot(page);
      expect(redone?.document).toEqual(after?.document);
      expect(redone?.selection).toEqual(after?.selection);
      expect(redone?.projectionJSON).toEqual(after?.projectionJSON);
    };

    for (const shortcut of ["Shift+Enter", `${MODIFIER}+Enter`]) {
      await reload();
      const before = await snapshot(page);
      await clearRefusals(page);
      await page.keyboard.press(shortcut);
      await expectNoRefusals(page);
      const inserted = await snapshot(page);
      expect(inserted?.document?.package.document.content).toHaveLength(1);
      const paragraph = inserted?.document?.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("Expected break paragraph.");
      const breaks = paragraph.content.flatMap((run) =>
        run.type === "run" ? run.content.filter((leaf) => leaf.type === "break") : [],
      );
      expect(breaks).toHaveLength(1);
      expect(breaks.at(0)).toMatchObject({
        type: "break",
        breakType: shortcut === "Shift+Enter" ? "textWrapping" : "page",
      });
      await undoExactly(before);
    }

    await reload();
    await page.keyboard.press("Enter");
    await select(page, 2);
    const split = await snapshot(page);
    await page.keyboard.press("Delete");
    expect((await snapshot(page))?.document?.package.document.content).toHaveLength(1);
    await undoExactly(split);

    await reload();
    const beforeShiftTab = await snapshot(page);
    await page.keyboard.press("Shift+Tab");
    expect(await snapshot(page)).toEqual(beforeShiftTab);

    await reload();
    await select(page, 1, 3);
    const beforeToolbar = await snapshot(page);
    await page.getByRole("button", { name: "Bold", exact: true }).click();
    const bold = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(
      bold?.type === "paragraph" &&
        bold.content.some((run) => run.type === "run" && run.formatting?.bold),
    ).toBe(true);
    await undoExactly(beforeToolbar);

    await reload();
    await select(page, 1, 3);
    await page.keyboard.press("Backspace");
    await page.keyboard.type("- ");
    await page.keyboard.press("Tab");
    const nested = await snapshot(page);
    await page.keyboard.press("Shift+Tab");
    const outdented = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(
      outdented?.type === "paragraph" &&
        outdented.formatting?.numPr?.kind === "reference" &&
        (outdented.formatting.numPr.ilvl ?? 0),
    ).toBe(0);
    await undoExactly(nested);
    await page.keyboard.type("item");
    await select(page, 1);
    const beforeListBackspace = await snapshot(page);
    await page.keyboard.press("Backspace");
    const removed = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(removed?.type === "paragraph" && removed.formatting?.numPr?.kind === "reference").toBe(
      false,
    );
    await undoExactly(beforeListBackspace);

    await reload();
    await select(page, 1, 3);
    await page.keyboard.press("Backspace");
    const beforeMarker = await snapshot(page);
    await page.keyboard.type("-");
    const beforeRule = await snapshot(page);
    await page.keyboard.type(" ");
    const ruled = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(ruled?.type === "paragraph" && ruled.formatting?.numPr?.kind).toBe("reference");
    await page.keyboard.press("Backspace");
    const ruleUndone = await snapshot(page);
    expect(ruleUndone?.document).toEqual(beforeRule?.document);
    expect(ruleUndone?.selection).toEqual(beforeRule?.selection);
    expect(ruleUndone?.text).toBe("-");
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeMarker?.document);
    expect((await snapshot(page))?.selection).toEqual(beforeMarker?.selection);

    await reload();
    await page.keyboard.type("x");
    const beforeFormatting = await snapshot(page);
    await select(page, 1, 4);
    const formattingSelection = await snapshot(page);
    await page.keyboard.press(`${MODIFIER}+i`);
    await select(page, 4);
    const beforeSecondTyping = await snapshot(page);
    await page.keyboard.type("y");
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeSecondTyping?.document);
    expect((await snapshot(page))?.selection).toEqual(beforeSecondTyping?.selection);
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeFormatting?.document);
    expect((await snapshot(page))?.selection).toEqual(formattingSelection?.selection);
  }
});

test(`canonical history replay ${CANONICAL_EVALUATION_HISTORY_REPLAY.seed} ${CANONICAL_EVALUATION_HISTORY_REPLAY.path}`, async ({
  page,
}) => {
  await runCanonicalHistoryReplay(page, CANONICAL_EVALUATION_HISTORY_REPLAY);
});

declare global {
  var __folioCollectPendingCanonicalLoad: (() => Promise<void>) | undefined;
  var __folioReportCollectedCanonicalLoad: (() => Promise<void>) | undefined;
  var __folioCompleteTransportLoad: (() => boolean) | undefined;
  var __folioTransportEvaluationStarted: (() => Promise<void>) | undefined;
  var __folioTransportPending:
    | WeakRef<Promise<boolean> & { complete: (loaded: boolean) => void }>
    | undefined;
}

const isProtocolRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

type CollectedTransportEvidenceOptions = {
  path: string;
  frameId: string;
};

const assertCollectedTransportEvidence = ({ path, frameId }: CollectedTransportEvidenceOptions) => {
  const events = readFileSync(path, "utf8")
    .trim()
    .split("\n")
    .map((line) => {
      const event: unknown = JSON.parse(line);
      if (!isProtocolRecord(event)) throw new TypeError("Malformed canonical protocol evidence");
      return event;
    });
  const sessions = new Set(
    events.flatMap((event) => {
      const params = event["params"];
      if (event["method"] !== "Runtime.executionContextCreated" || !isProtocolRecord(params))
        return [];
      const context = params["context"];
      if (!isProtocolRecord(context)) return [];
      const auxiliary = context["auxData"];
      const sessionId = event["sessionId"];
      return isProtocolRecord(auxiliary) &&
        auxiliary["frameId"] === frameId &&
        typeof sessionId === "string"
        ? [sessionId]
        : [];
    }),
  );
  expect(sessions.size, "protocol evidence must identify this test frame").toBeGreaterThan(0);
  const errors = events.filter((event) => {
    const sessionId = event["sessionId"];
    return typeof sessionId === "string" && sessions.has(sessionId) && event["error"] !== undefined;
  });
  expect(errors, "both unrooted calls must produce correlated raw CDP failures").toHaveLength(2);
  const methods = [];
  for (const error of errors) {
    expect(error["error"]).toEqual({ code: -32000, message: "Promise was collected" });
    const requests = events.filter(
      (event) =>
        event["id"] === error["id"] &&
        event["sessionId"] === error["sessionId"] &&
        event["method"] !== undefined,
    );
    expect(requests).toHaveLength(1);
    expect(requests.at(0)).toMatchObject({ awaitPromise: true });
    methods.push(requests.at(0)?.["method"]);
  }
  expect(methods.sort()).toEqual(["Runtime.callFunctionOn", "Runtime.evaluate"]);
};

test("canonical load roots the awaited transport promise across forced collection", async ({
  page,
}) => {
  const protocolPath = process.env["FOLIO_CANONICAL_PROTOCOL_LOG"];
  if (!protocolPath)
    throw new TypeError(
      "Run this transport regression through canonical-interactions.ts to capture raw CDP evidence",
    );
  await page.goto("/?session=canonical");
  await waitForCanonicalPageReady(page);
  const cdp = await page.context().newCDPSession(page);
  let entered = Promise.withResolvers<void>();
  await page.exposeFunction("__folioTransportEvaluationStarted", () => entered.resolve());
  try {
    await cdp.send("Runtime.enable");
    const { frameTree } = await cdp.send("Page.getFrameTree");
    for (const transport of ["unrooted", "bridge", "handle", "handle-released", "cdp"] as const) {
      entered = Promise.withResolvers<void>();
      await page.evaluate((fixtureTransport) => {
        const bridge = globalThis.__folioCanonical;
        if (!bridge) throw new TypeError("Canonical bridge unavailable");
        const createOperation = () => {
          const { promise, resolve } = Promise.withResolvers<boolean>();
          return Object.assign(promise, { complete: resolve });
        };
        globalThis.__folioCompleteTransportLoad = () => {
          const pending = globalThis.__folioTransportPending?.deref();
          if (!pending) return false;
          pending.complete(true);
          return true;
        };
        bridge.load = () => {
          const pending = createOperation();
          globalThis.__folioTransportPending = new WeakRef(pending);
          // Only this case retains a source promise through the bridge closure.
          if (fixtureTransport === "bridge") bridge.load = () => pending;
          void globalThis.__folioTransportEvaluationStarted?.();
          return pending;
        };
      }, transport);

      if (transport === "handle" || transport === "handle-released") {
        const operation = await page.evaluateHandle(
          "({ pending: globalThis.__folioCanonical.load([]) })",
        );
        try {
          // Both cases use the identical evaluator; releasing the handle is the ablation.
          const loaded = Result.tryPromise({
            try: () => operation.evaluate(({ pending }) => pending),
            catch: (error: unknown) => error,
          });
          // A round trip on the evaluator's session establishes the call was sent
          // before the remote root is released (no timing or retry barrier).
          await operation.evaluate(() => true);
          if (transport === "handle-released") await operation.dispose();
          await cdp.send("HeapProfiler.collectGarbage");
          const retained = await page.evaluate(() => globalThis.__folioCompleteTransportLoad?.());
          const result = await loaded;
          if (transport === "handle") {
            expect(retained).toBe(true);
            expect(result.isOk()).toBe(true);
            if (result.isOk()) expect(result.value).toBe(true);
          } else {
            expect(retained).toBe(false);
            expect(result.isErr()).toBe(true);
            if (result.isErr())
              expect(String(result.error)).toContain("Execution context was destroyed");
          }
        } finally {
          await operation.dispose();
        }
        continue;
      }

      const loaded = Result.tryPromise({
        try: () =>
          transport === "cdp"
            ? evaluateCanonicalLoad(page, [])
            : cdp
                .send("Runtime.evaluate", {
                  expression: "globalThis.__folioCanonical.load([])",
                  awaitPromise: true,
                  returnByValue: true,
                })
                .then(({ result }) => result.value),
        catch: (error: unknown) => error,
      });
      await entered.promise;
      // One forced collection, after the load has returned across a protocol boundary.
      await cdp.send("HeapProfiler.collectGarbage");
      const retained = await page.evaluate(() => globalThis.__folioCompleteTransportLoad?.());
      const result = await loaded;
      if (transport === "unrooted") {
        expect(retained).toBe(false);
        expect(result.isErr()).toBe(true);
        if (result.isErr()) expect(String(result.error)).toContain("Promise was collected");
        continue;
      }
      expect(retained).toBe(true);
      expect(result.isOk()).toBe(true);
      if (result.isOk()) expect(result.value).toBe(true);
    }
    assertCollectedTransportEvidence({ path: protocolPath, frameId: frameTree.frame.id });
  } finally {
    const evidence = readFileSync(protocolPath, "utf8");
    console.log(`Canonical transport protocol evidence:\n${evidence}`);
    await test.info().attach("canonical-transport-protocol", {
      body: evidence,
      contentType: "application/x-ndjson",
    });
    await cdp.detach();
  }
});

test("canonical history load survives browser collection while its evaluation is pending", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await waitForCanonicalPageReady(page);
  const cdp = await page.context().newCDPSession(page);
  const collections: Promise<void>[] = [];
  const lostOwner = Promise.withResolvers<never>();
  await page.exposeFunction("__folioReportCollectedCanonicalLoad", () => {
    lostOwner.reject(new TypeError("Canonical load promise lost its transport owner"));
  });
  await page.exposeFunction("__folioCollectPendingCanonicalLoad", () => {
    const collect = async () => {
      for (let collection = 0; collection < 20; collection++)
        await cdp.send("HeapProfiler.collectGarbage");
    };
    const pending = collect();
    collections.push(pending);
    return pending;
  });
  await page.evaluate(() => {
    const bridge = globalThis.__folioCanonical;
    if (!bridge) throw new TypeError("Canonical bridge unavailable");
    const collect = globalThis.__folioCollectPendingCanonicalLoad;
    if (!collect) throw new TypeError("Canonical collection probe unavailable");
    const report = globalThis.__folioReportCollectedCanonicalLoad;
    if (!report) throw new TypeError("Canonical collection observer unavailable");
    const load = bridge.load;
    const createOperation = () => {
      const completion = Promise.withResolvers<boolean>();
      return Object.assign(completion.promise, {
        complete: completion.resolve,
        fail: completion.reject,
      });
    };
    bridge.load = (bytes) => {
      const pending = createOperation();
      const weak = new WeakRef(pending);
      // The transport must own its returned promise. A weak completion link
      // lets collection exercise that boundary after the real load commits,
      // independently of incidental roots inside the parser or adapter.
      load(bytes)
        .then(
          (loaded) =>
            collect().then(() => {
              const owner = weak.deref();
              if (!owner) return report();
              return owner.complete(loaded);
            }),
          (error: unknown) => weak.deref()?.fail(error),
        )
        .catch((error: unknown) => weak.deref()?.fail(error));
      return pending;
    };
  });
  const runCases = async () => {
    const source = [
      ...new Uint8Array(await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }))),
    ];
    const loadedFixture = await canonicalLoadFixture(
      await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" })),
    );
    // Exercise the structural helper too, including reloads after real history
    // edits. The original collection probe covered only the history oracle.
    for (let reload = 0; reload < 3; reload++) {
      await loadReady(page, loadedFixture);
      await select(page, 2);
      const before = await snapshot(page);
      await clearRefusals(page);
      await page.keyboard.press("Shift+Enter");
      const edited = await snapshot(page);
      expect(edited?.document).not.toEqual(before?.document);
      await page.keyboard.press(`${MODIFIER}+z`);
      expect((await snapshot(page))?.document).toEqual(before?.document);
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      expect((await snapshot(page))?.document).toEqual(edited?.document);
      await expectNoRefusals(page);
    }
    await initializeCanonicalBrowserHistory(page, source);
    const actions = canonicalHistoryReplayActions(CANONICAL_EVALUATION_HISTORY_REPLAY);
    expect(
      await checkCanonicalBrowserHistory({
        page,
        source,
        actions,
        missing: createMissingOpBurndown(),
      }),
    ).toBeGreaterThan(0);
    await Promise.all(collections);
    expect(
      collections.length,
      "collection must exercise the actual canonical load boundary",
    ).toBeGreaterThan(0);
  };
  try {
    await Promise.race([runCases(), lostOwner.promise]);
  } finally {
    await cdp.detach();
  }
});
