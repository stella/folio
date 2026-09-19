/**
 * The pipeline-totality invariant, over a package built in memory.
 *
 * Never a corpus file: the corpus is fetched, large and licence-bound, and a
 * test that needed one would neither run on a fresh clone nor say anything
 * about the invariant that a one-paragraph package cannot say. What is worth
 * asserting here is the machinery — every stage runs, a throwing stage is
 * reported without stopping the stages after it — and the one substantive
 * claim that does not need a hard document: comparing a package with itself
 * reports nothing.
 */

import { describe, expect, test } from "bun:test";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import type { CorpusInvariantInput } from "./lib/corpus-invariants/contract";
import { DEFAULT_INVARIANT_BUDGET_MS } from "./lib/corpus-invariants/contract";
import {
  PIPELINE_STAGE_TABLE,
  PIPELINE_STAGES,
  type PipelineStage,
  runPipelineStages,
  runPipelineTotalityInvariant,
} from "./lib/corpus-invariants/pipeline-totality";

const minimalInput = async (): Promise<CorpusInvariantInput> => {
  const buffer = await createDocx(createEmptyDocument({ initialText: "One paragraph." }));
  return {
    bytes: new Uint8Array(buffer),
    buffer,
    parsed: await parseDocx(buffer, { preloadFonts: false }),
    documentPart: "word/document.xml",
    budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
  };
};

describe("runPipelineTotalityInvariant", () => {
  test("a one-paragraph package survives every published entry point", async () => {
    const { failures, timings } = await runPipelineTotalityInvariant(await minimalInput());

    expect(failures).toEqual([]);
    expect(Object.keys(timings).sort()).toEqual([...PIPELINE_STAGES].sort());
  });

  test("comparing the package with itself reports no changes", async () => {
    const compareSelf = PIPELINE_STAGE_TABLE.find(({ name }) => name === "compare-self");
    if (compareSelf === undefined) {
      throw new Error("the stage table lost compare-self");
    }

    const problems = await compareSelf.run({ input: await minimalInput(), laidOut: null });

    expect(problems).toEqual([]);
  });
});

describe("runPipelineStages", () => {
  const recordingStage = (name: string, ran: string[]): PipelineStage => ({
    name,
    run: () => {
      ran.push(name);
      return Promise.resolve([]);
    },
  });

  test("a throwing stage is reported and the stages after it still run", async () => {
    const ran: string[] = [];
    const stages = [
      recordingStage("before", ran),
      {
        name: "boom",
        run: () => Promise.reject(new TypeError("layout walked off the page")),
      },
      recordingStage("after", ran),
    ] as const satisfies readonly PipelineStage[];

    const { failures, timings } = await runPipelineStages(stages, await minimalInput());

    expect(ran).toEqual(["before", "after"]);
    expect(failures).toHaveLength(1);
    expect(failures[0]?.message).toBe("boom: TypeError: layout walked off the page");
    expect(Object.keys(timings).sort()).toEqual(["after", "before", "boom"]);
  });

  test("two stages failing the same way stay two signatures", async () => {
    const throwing = (name: string): PipelineStage => ({
      name,
      run: () => Promise.reject(new TypeError("the same defect")),
    });

    const { failures } = await runPipelineStages(
      [throwing("first"), throwing("second")],
      await minimalInput(),
    );

    expect(failures.map(({ message }) => message)).toEqual([
      "first: TypeError: the same defect",
      "second: TypeError: the same defect",
    ]);
  });

  test("an assertion carries its stage name and no per-file particulars", async () => {
    const stages = [
      {
        name: "asserting",
        run: () => Promise.resolve(["comparing a package with itself reported changes"]),
      },
    ] as const satisfies readonly PipelineStage[];

    const { failures } = await runPipelineStages(stages, await minimalInput());

    expect(failures).toEqual([
      {
        invariant: "pipeline-totality",
        message: "asserting: comparing a package with itself reported changes",
        frame: "-",
      },
    ]);
  });
});
