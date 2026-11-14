/**
 * Typing latency in long documents, report-only.
 *
 * Opens generated documents in the playground, types at a caret in the middle
 * of the document (track changes on; `FOLIO_TYPING_LATENCY_MODE=editing` turns
 * them off), and logs one JSON line per run, for each engine and painter:
 *
 * - `inputMs`: main-thread time the browser spent in the key's input events
 *   (keydown, keypress, beforeinput), where the editor commits the edit. Read
 *   from the Event Timing API, so it holds no time spent waiting.
 * - `frameMs`: from keydown to the end of the first task after the next
 *   animation frame, so the work the edit schedules (layout, paint) counts.
 * - `measuredBlocksPerKey` and `layoutsPerKey`: what layout did per key, read
 *   through the layout instrumentation hook.
 * - `commitsPerKey` (React host; `null` for a Vue host): React commits per
 *   key that change the page, counted through the developer tools hook. Each one with a focused
 *   editor reads the editor's whole DOM to save the selection, so commits
 *   cost on their own.
 *   `FOLIO_TYPING_LATENCY_COMMITS=1` also logs which components' state
 *   changed in those commits.
 *
 * It asserts nothing: the numbers depend on the machine. Run it with
 * `bunx playwright test --project=typing-latency`, add `--headed` for a
 * headed browser, and set `FOLIO_TYPING_LATENCY_PROFILE=<dir>` to also write a
 * CPU profile of the measured keys for each run.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import { type Page, test } from "../support/unclassifiedProbe";

const WARM_KEYS = 5;
const MEASURED_KEYS = Number(process.env["FOLIO_TYPING_LATENCY_KEYS"] ?? 30);
const PAUSE_MS = 150;
const PROFILE_DIRECTORY = process.env["FOLIO_TYPING_LATENCY_PROFILE"];
const MODE =
  process.env["FOLIO_TYPING_LATENCY_MODE"] === "editing"
    ? "editing"
    : "suggesting";

const WORDS =
  "the supplier shall deliver goods services under this agreement subject to terms conditions warranty liability indemnity confidential information party notice period".split(
    " ",
  );

/** A contract of numbered articles, clauses, short lists and fee tables. */
const contract = (articles: number): string => {
  let seed = 7;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31;
    return seed / 2 ** 31;
  };
  const sentence = (length: number) =>
    Array.from(
      { length },
      () => WORDS[Math.floor(random() * WORDS.length)] ?? "the",
    ).join(" ");
  const lines = ["# Services Agreement", ""];
  for (let article = 1; article <= articles; article += 1) {
    lines.push(`## ${article}. ${sentence(3)}`, "");
    for (let clause = 1; clause <= 8; clause += 1) {
      lines.push(
        `${article}.${clause} **${sentence(2)}** ${sentence(40)}. ${sentence(25)}.`,
        "",
      );
      if (clause % 3 === 0) {
        for (let item = 1; item <= 4; item += 1)
          lines.push(`${item}. ${sentence(18)};`);
        lines.push("");
      }
    }
    if (article % 10 === 0) {
      lines.push("| Item | Fee | Term |", "|---|---|---|");
      for (let row = 0; row < 12; row += 1)
        lines.push(`| ${sentence(3)} | ${row * 100} | ${sentence(2)} |`);
      lines.push("");
    }
  }
  return lines.join("\n");
};

/** One numbered list of short items. */
const numberedList = (items: number): string =>
  Array.from({ length: items }, (_, index) => `1. Item ${index + 1}`).join(
    "\n",
  );

/** Plain paragraphs only. */
const plainParagraphs = (paragraphs: number): string =>
  Array.from(
    { length: paragraphs },
    (_, index) => `Paragraph ${index + 1} holds a sentence of plain text.`,
  ).join("\n\n");

const DOCUMENTS = {
  "plain paragraphs, 1,500": () => plainParagraphs(1500),
  "contract, 40 articles": () => contract(40),
  "contract, 100 articles": () => contract(100),
  "numbered list, 3,500 items": () => numberedList(3500),
} as const;

const ENGINES = ["default", "canonical"] as const;
const PAINTERS = ["stock", "display-list"] as const;
/** The React playground serves every painter; the Vue one the stock painter. */
const HOSTS = [
  { name: "react", origin: "", painters: PAINTERS },
  { name: "vue", origin: PLAYGROUND_HOSTS.vue, painters: ["stock"] },
] as const;

test.use({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 });

const load = async (page: Page, bytes: number[]) => {
  await page.evaluate(async (data) => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!ref) throw new TypeError("The playground editor is not ready.");
    await ref.loadDocumentBuffer(new Uint8Array(data));
    // The default engine creates its editor view on first use.
    ref.ensureEditorView();
  }, bytes);
  await page.waitForFunction(
    () =>
      (globalThis.__folioPlayground?.getEditorRef()?.getEditor()?.getView()
        ?.state.doc.childCount ?? 0) > 500,
    undefined,
    { timeout: 120_000 },
  );
  // Let the first layout of every page finish before typing.
  await page.waitForTimeout(5000);
};

/**
 * Bring the page of the middle paragraph into view first: a host may find a
 * paragraph to scroll to only on a page it has painted.
 */
const showMiddlePage = async (page: Page) => {
  await page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const view = ref?.getEditor()?.getView();
    if (!ref || !view) throw new TypeError("No editor view.");
    const paragraphs: number[] = [];
    view.state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.textContent.length > 4)
        paragraphs.push(position);
      return true;
    });
    const middle = paragraphs[Math.floor(paragraphs.length / 2)];
    if (middle === undefined) throw new TypeError("No paragraph to type in.");
    const pageNumber = ref.getEditorRef()?.getPageNumberForPmPos(middle + 3);
    if (pageNumber === undefined || pageNumber === null)
      throw new TypeError("No page holds the paragraph.");
    ref.scrollToPage(pageNumber);
  });
  await page.waitForTimeout(1000);
};

/** Put the caret inside a paragraph in the middle and start recording. */
const prepare = (page: Page) =>
  page.evaluate(() => {
    const editor = globalThis.__folioPlayground?.getEditorRef()?.getEditor();
    const view = editor?.getView();
    if (!editor || !view) throw new TypeError("No editor view.");
    const paragraphs: number[] = [];
    view.state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.textContent.length > 4)
        paragraphs.push(position);
      return true;
    });
    const middle = paragraphs[Math.floor(paragraphs.length / 2)];
    if (middle === undefined) throw new TypeError("No paragraph to type in.");
    editor.setSelection(middle + 3, middle + 3);
    // Type where the caret is in view, as a person does: the edited page and
    // the caret are painted.
    const paraId: unknown = view.state.doc.nodeAt(middle)?.attrs["paraId"];
    if (
      typeof paraId !== "string" ||
      !globalThis.__folioPlayground?.getEditorRef()?.scrollToParaId(paraId)
    )
      throw new TypeError("No paragraph to scroll to.");
    editor.focus();
    const record = { input: new Map<number, number>(), frames: [] as number[] };
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries() as PerformanceEventTiming[]) {
        if (entry.interactionId === 0) continue;
        record.input.set(
          entry.interactionId,
          (record.input.get(entry.interactionId) ?? 0) +
            entry.processingEnd -
            entry.processingStart,
        );
      }
    }).observe({ type: "event", durationThreshold: 16 });
    window.addEventListener(
      "keydown",
      () => {
        const start = performance.now();
        requestAnimationFrame(() => {
          const channel = new MessageChannel();
          channel.port1.onmessage = () =>
            record.frames.push(performance.now() - start);
          channel.port2.postMessage(undefined);
        });
      },
      { capture: true },
    );
    (globalThis as { __typingLatency?: typeof record }).__typingLatency =
      record;
  });

const readLayoutCounts = (page: Page) =>
  page.evaluate(
    () =>
      (
        globalThis as {
          __typingLayoutCounts?: { measuredBlocks: number; layouts: number };
        }
      ).__typingLayoutCounts ?? { measuredBlocks: 0, layouts: 0 },
  );

const readCommits = (page: Page) =>
  page.evaluate(() => {
    const commits = (
      globalThis as {
        __typingCommits?: {
          count: number;
          withEffects: number;
          changed: Record<string, number>;
        };
      }
    ).__typingCommits;
    return {
      count: commits?.count ?? 0,
      withEffects: commits?.withEffects ?? 0,
      changed: { ...commits?.changed },
    };
  });

const median = (values: readonly number[]): number | null => {
  const sorted = values.toSorted((left, right) => left - right);
  return sorted.length === 0 ? null : (sorted[sorted.length >> 1] ?? null);
};

for (const [name, markdown] of Object.entries(DOCUMENTS)) {
  for (const engine of ENGINES) {
    for (const host of HOSTS) {
      for (const painter of host.painters) {
        const hostSuffix = host.name === "react" ? "" : `, ${host.name} host`;
        test(`typing latency: ${name}, ${engine} engine, ${painter} painter${hostSuffix}`, async ({
          page,
        }) => {
          test.setTimeout(300_000);
          const bytes = [
            ...new Uint8Array(await createDocx(fromMarkdown(markdown()))),
          ];
          const query = new URLSearchParams({ mode: MODE });
          if (engine === "canonical") query.set("session", "canonical");
          if (painter === "display-list") query.set("pageRenderer", painter);
          // Counts what each layout pass measured, through the layout
          // instrumentation hook the editor already calls.
          await page.addInitScript(() => {
            const commits = {
              count: 0,
              withEffects: 0,
              changed: {} as Record<string, number>,
            };
            (
              globalThis as { __typingCommits?: typeof commits }
            ).__typingCommits = commits;
            type Hook = {
              memoizedState: unknown;
              queue: unknown;
              next: Hook | null;
            };
            type Fiber = {
              tag: number;
              type: unknown;
              alternate: Fiber | null;
              memoizedState: unknown;
              child: Fiber | null;
              sibling: Fiber | null;
              flags: number;
              subtreeFlags: number;
            };
            const nameOf = (fiber: Fiber): string => {
              type Named = {
                displayName?: string;
                name?: string;
                render?: Named;
                type?: Named;
              } | null;
              const type = fiber.type as Named;
              const inner = type?.render ?? type?.type ?? type;
              return (
                type?.displayName ??
                inner?.displayName ??
                inner?.name ??
                type?.name ??
                "anonymous"
              );
            };
            /** Which fields of a changed state value changed, to name it. */
            const describeChange = (
              value: unknown,
              before: unknown,
            ): string => {
              if (value instanceof Element)
                return `<${value.tagName} ${value.getAttribute("aria-label") ?? ""}>`;
              if (value === null || typeof value !== "object")
                return JSON.stringify(value)?.slice(0, 40) ?? "";
              if (Array.isArray(value)) return `[${String(value.length)}]`;
              const fields = Object.keys(value).filter(
                (field) =>
                  before === null ||
                  typeof before !== "object" ||
                  !Object.is(
                    Reflect.get(value, field),
                    Reflect.get(before, field),
                  ),
              );
              return `{${fields.slice(0, 6).join(",")}}`;
            };
            const collect = (fiber: Fiber | null): void => {
              for (let current = fiber; current; current = current.sibling) {
                const previous = current.alternate;
                // Function components and their forwardRef / memo forms.
                // Components that rendered in this commit (PerformedWork); a
                // fiber that did not keeps an older alternate.
                if (
                  previous &&
                  (current.flags & 1) !== 0 &&
                  [0, 11, 14, 15].includes(current.tag)
                ) {
                  let hook = current.memoizedState as Hook | null;
                  let old = previous.memoizedState as Hook | null;
                  let index = 0;
                  while (hook && old) {
                    if (
                      hook.queue !== null &&
                      hook.memoizedState !== old.memoizedState
                    ) {
                      const shape = describeChange(
                        hook.memoizedState,
                        old.memoizedState,
                      );
                      const key = `${nameOf(current)}#${String(index)} ${shape ?? ""}`;
                      commits.changed[key] = (commits.changed[key] ?? 0) + 1;
                    }
                    hook = hook.next;
                    old = old.next;
                    index += 1;
                  }
                }
                collect(current.child);
              }
            };
            (
              globalThis as { __REACT_DEVTOOLS_GLOBAL_HOOK__?: unknown }
            ).__REACT_DEVTOOLS_GLOBAL_HOOK__ = {
              supportsFiber: true,
              renderers: new Map(),
              inject: () => 1,
              onScheduleFiberRoot: () => undefined,
              onCommitFiberRoot: (_id: number, root: { current: Fiber }) => {
                commits.count += 1;
                // React's before-mutation, mutation, layout and passive masks: a
                // commit without them skips the selection save.
                if ((root.current.subtreeFlags | root.current.flags) & 13878)
                  commits.withEffects += 1;
                collect(root.current);
              },
              onCommitFiberUnmount: () => undefined,
              onPostCommitFiberRoot: () => undefined,
              checkDCE: () => undefined,
            };
            const counts = { measuredBlocks: 0, layouts: 0 };
            (
              globalThis as { __typingLayoutCounts?: typeof counts }
            ).__typingLayoutCounts = counts;
            globalThis.__folioLayoutInstrumentation = {
              onLayoutComplete: () => {
                counts.layouts += 1;
              },
              onMeasureBlock: () => {
                counts.measuredBlocks += 1;
              },
            };
          });
          await page.goto(`${host.origin}/?${query.toString()}`);
          await page.waitForSelector(".layout-page");
          await load(page, bytes);
          await showMiddlePage(page);
          await prepare(page);
          for (let key = 0; key < WARM_KEYS; key += 1) {
            await page.keyboard.press("x");
            await page.waitForTimeout(PAUSE_MS);
          }
          const cdp =
            PROFILE_DIRECTORY === undefined
              ? undefined
              : await page.context().newCDPSession(page);
          if (cdp) {
            await cdp.send("Profiler.enable");
            await cdp.send("Profiler.setSamplingInterval", { interval: 200 });
            await cdp.send("Profiler.start");
          }
          const skip = await page.evaluate(
            () =>
              (globalThis as { __typingLatency?: { frames: number[] } })
                .__typingLatency?.frames.length ?? 0,
          );
          const countsBefore = await readLayoutCounts(page);
          const commitsBefore = await readCommits(page);
          for (let key = 0; key < MEASURED_KEYS; key += 1) {
            await page.keyboard.press("x");
            await page.waitForTimeout(PAUSE_MS);
          }
          if (cdp && PROFILE_DIRECTORY !== undefined) {
            const { profile } = await cdp.send("Profiler.stop");
            mkdirSync(PROFILE_DIRECTORY, { recursive: true });
            writeFileSync(
              join(
                PROFILE_DIRECTORY,
                `${name.replaceAll(/[^a-z0-9]+/gu, "-")}-${host.name}-${engine}-${painter}.cpuprofile`,
              ),
              JSON.stringify(profile),
            );
          }
          const measured = await page.evaluate((from) => {
            const record = (
              globalThis as {
                __typingLatency?: {
                  input: Map<number, number>;
                  frames: number[];
                };
              }
            ).__typingLatency;
            return {
              input: [...(record?.input.values() ?? [])].slice(from),
              frames: record?.frames.slice(from) ?? [],
            };
          }, skip);
          const countsAfter = await readLayoutCounts(page);
          const commitsAfter = await readCommits(page);
          console.info(
            `typing latency ${JSON.stringify({
              document: name,
              host: host.name,
              engine,
              painter,
              mode: MODE,
              keys: measured.frames.length,
              inputMs: median(measured.input),
              frameMs: median(measured.frames),
              layoutsPerKey:
                (countsAfter.layouts - countsBefore.layouts) / MEASURED_KEYS,
              measuredBlocksPerKey:
                (countsAfter.measuredBlocks - countsBefore.measuredBlocks) /
                MEASURED_KEYS,
              // React commits only: a Vue host has none to count.
              commitsPerKey:
                host.name === "react"
                  ? (commitsAfter.withEffects - commitsBefore.withEffects) /
                    MEASURED_KEYS
                  : null,
            })}`,
          );
          if (
            host.name === "react" &&
            process.env["FOLIO_TYPING_LATENCY_COMMITS"] === "1"
          )
            console.info(
              `typing latency state changes ${JSON.stringify(
                Object.fromEntries(
                  Object.entries(commitsAfter.changed)
                    .map(
                      ([key, count]) =>
                        [
                          key,
                          (count - (commitsBefore.changed[key] ?? 0)) /
                            MEASURED_KEYS,
                        ] as const,
                    )
                    .filter(([, perKey]) => perKey > 0),
                ),
              )}`,
            );
        });
      }
    }
  }
}
