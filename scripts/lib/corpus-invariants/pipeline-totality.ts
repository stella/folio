/**
 * Every entry point folio publishes must survive every package that parses.
 *
 * The gate's other invariants all run down one road: parse, edit, save, parse
 * again. That road proves the DOCX round trip and nothing else. A package that
 * round-trips perfectly can still abort the layout engine on a section it has
 * no measure for, walk off the end of a display list, write a PDF that never
 * finishes, or throw on the markdown projection, and none of those defects
 * would be visible to an invariant that only re-serializes. They are the
 * defects a user meets first, because the editor lays out before it saves.
 *
 * So this invariant runs the downstream entry points instead, over a package
 * that has already proven it parses. Nothing here asserts that any stage
 * produces the right answer: a display list with a wrong glyph is a different
 * invariant's problem. What is asserted is totality, the weakest claim worth
 * making and the one that fails loudest: no published entry point throws on a
 * package folio accepted.
 *
 * Stages are independent by construction. A stage that throws is recorded and
 * the next stage still runs, because a package that breaks layout usually
 * breaks something else too, and learning that in one pass costs one parse
 * instead of six. Each stage's failure carries the stage name in its message,
 * so two stages dying of the same error stay two signatures.
 */

import { compareDocx } from "@stll/folio-core/compare/compare";
import { buildDisplayList } from "@stll/folio-core/display-list/build/buildDisplayList";
import { exportDocxToPdf } from "@stll/folio-core/export-pdf";
import {
  type HeadlessFontSource,
  installHeadlessMeasureProvider,
} from "@stll/folio-core/fonts/headlessMeasure";
import { type HeadlessLayoutResult, layoutDocxHeadless } from "@stll/folio-core/headless-layout";
import { toMarkdown } from "@stll/folio-core/markdown";
import { FolioDocxReviewer } from "@stll/folio-core/server";
import { Result } from "better-result";

import {
  type CorpusFailure,
  describeError,
  failureFromAssertion,
  failureFromError,
} from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  type StageTimings,
  timeStage,
} from "./contract";

const STAGE_NAMES = {
  layout: "layout",
  displayList: "display-list",
  pdf: "pdf",
  markdown: "markdown",
  agentsSnapshot: "agents-snapshot",
  compareSelf: "compare-self",
} as const;

/**
 * An instant, rather than the clock, everywhere a stage stamps one.
 *
 * A PDF's creation date and a tracked change's timestamp both reach the output
 * bytes. Reading them from the clock would make a stage's result differ between
 * two runs over one package, which is the opposite of what a gate wants.
 */
const FIXED_TIMESTAMP = "2026-01-01T00:00:00.000Z";

const COMPARE_AUTHOR = "corpus-gate";

/** {@link CompareVerification} and {@link CompareCompatibility} when nothing is wrong. */
const VERIFIED = "verified";
const STANDARD_OOXML = "standard-ooxml";

/**
 * No face binaries at all.
 *
 * The provider measures every face with its documented stand-in, which lays out
 * pages nobody should trust and is exactly right here: this invariant asks
 * whether the engines run, not where the lines break. Shipping a font set would
 * make the gate's verdict depend on which fonts the machine had.
 */
const NO_FONTS = {
  load: () => [],
} as const satisfies HeadlessFontSource;

/**
 * Whether the process-wide measure provider is installed.
 *
 * Installed lazily and once, never at import: the provider is global state, and
 * a module that installed one on import would silently change how anything else
 * in the process measures. `exportDocxToPdf` installs its own for the duration
 * of an export and restores whatever it found, so the order of the stages does
 * not matter.
 */
let measureProviderInstalled = false;

const ensureMeasureProvider = (): void => {
  if (measureProviderInstalled) {
    return;
  }
  installHeadlessMeasureProvider(NO_FONTS);
  measureProviderInstalled = true;
};

/**
 * What the stages share: the file, and whatever an earlier stage produced that
 * a later one needs. `laidOut` is the only such hand-off, and it is null until
 * `layout` has run and succeeded.
 */
export type PipelineContext = {
  readonly input: CorpusInvariantInput;
  laidOut: HeadlessLayoutResult | null;
};

/**
 * One entry point under test.
 *
 * `run` reports assertion messages, one per problem, and throws nothing on
 * purpose: a throw out of `run` is the defect this invariant exists to catch,
 * and the runner turns it into a failure that names the throwing folio frame.
 */
export type PipelineStage = {
  readonly name: string;
  readonly run: (context: PipelineContext) => Promise<readonly string[]>;
};

export const PIPELINE_STAGE_TABLE = [
  {
    name: STAGE_NAMES.layout,
    run: async (context) => {
      ensureMeasureProvider();
      const laidOut = await layoutDocxHeadless(context.input.buffer);
      if (laidOut.isErr()) {
        return [describeError(laidOut.error)];
      }
      context.laidOut = laidOut.value;
      return [];
    },
  },
  {
    name: STAGE_NAMES.displayList,
    run: ({ laidOut }) => {
      // Nothing to build when layout never produced a result; the layout stage
      // has already reported why, and a second failure would say no more.
      if (laidOut !== null) {
        buildDisplayList({ layout: laidOut.layout, blockLookup: laidOut.blockLookup });
      }
      return Promise.resolve([]);
    },
  },
  {
    name: STAGE_NAMES.pdf,
    run: async ({ input }) => {
      ensureMeasureProvider();
      const exported = await exportDocxToPdf(input.buffer, {
        fonts: NO_FONTS,
        timestamp: FIXED_TIMESTAMP,
      });
      return exported.isErr() ? [describeError(exported.error)] : [];
    },
  },
  {
    name: STAGE_NAMES.markdown,
    run: ({ input }) => {
      toMarkdown(input.parsed);
      return Promise.resolve([]);
    },
  },
  {
    name: STAGE_NAMES.agentsSnapshot,
    run: async ({ input }) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(input.buffer);
      reviewer.snapshot();
      return [];
    },
  },
  {
    name: STAGE_NAMES.compareSelf,
    run: async ({ input }) => {
      // Two independent copies: the comparison holds both sides at once, and a
      // single buffer passed twice would let either side's handling alias the
      // other's bytes.
      const compared = await compareDocx(input.buffer.slice(0), input.buffer.slice(0), {
        author: COMPARE_AUTHOR,
        timestamp: FIXED_TIMESTAMP,
      });
      if (compared.isErr()) {
        return [describeError(compared.error)];
      }
      const { changes, compatibility, verification } = compared.value;
      const problems: string[] = [];
      if (changes.length > 0) {
        problems.push("comparing a package with itself reported changes");
      }
      if (verification.status !== VERIFIED) {
        problems.push(`the self-comparison is ${verification.status}`);
      }
      if (compatibility.status !== STANDARD_OOXML) {
        problems.push(`the self-comparison encodes revisions as ${compatibility.status}`);
      }
      return problems;
    },
  },
] as const satisfies readonly PipelineStage[];

/** The stage names, for the census's slowest-files report. */
export const PIPELINE_STAGES: readonly string[] = PIPELINE_STAGE_TABLE.map(({ name }) => name);

const failureFromStageError = (stage: string, cause: unknown): CorpusFailure => {
  const failure = failureFromError(EXTENDED_CORPUS_INVARIANTS.pipelineTotality, cause);
  return { ...failure, message: `${stage}: ${failure.message}` };
};

/**
 * Run a stage table end to end, one stage at a time.
 *
 * Exported over an arbitrary table rather than closed over the real one so the
 * independence claim is testable: a table with a stage that throws in the
 * middle must still report every stage after it.
 */
export const runPipelineStages = async (
  stages: readonly PipelineStage[],
  input: CorpusInvariantInput,
): Promise<CorpusInvariantOutcome> => {
  const timings: StageTimings = {};
  const failures: CorpusFailure[] = [];
  const context: PipelineContext = { input, laidOut: null };

  for (const { name, run } of stages) {
    // oxlint-disable-next-line no-await-in-loop -- the stages share the process-wide
    // measure provider and hand results forward, so they cannot overlap.
    const outcome = await timeStage(timings, name, () =>
      Result.tryPromise({ try: () => run(context), catch: (cause: unknown) => cause }),
    );
    if (outcome.isErr()) {
      failures.push(failureFromStageError(name, outcome.error));
      continue;
    }
    for (const problem of outcome.value) {
      failures.push(
        failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.pipelineTotality, `${name}: ${problem}`),
      );
    }
  }

  return { failures, timings };
};

export const runPipelineTotalityInvariant = (
  input: CorpusInvariantInput,
): Promise<CorpusInvariantOutcome> => runPipelineStages(PIPELINE_STAGE_TABLE, input);
