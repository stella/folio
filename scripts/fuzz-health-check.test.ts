import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";

import { propertyConfig, propertyTestTimeout } from "../test/property-testing";

import { HEALTH_MARKER } from "../test/consumer-scenarios/support/fuzz-health";
import { classifyFuzzRun } from "../test/fuzz-health";
import { parseConsumerArgs } from "./consumer-scenario-args";
import { checkFuzzHealth } from "./fuzz-health-check";

setDefaultTimeout(propertyTestTimeout(5_000));

const line = (value: unknown) => `${HEALTH_MARKER}${JSON.stringify(value)}`;

describe("fuzz infrastructure health", () => {
  test("Bun's actual leading separator argv still selects the requested scenario", async () => {
    const process = Bun.spawn(
      [
        Bun.which("bun") ?? panic("Bun executable unavailable"),
        "scripts/consumer-scenario-args.ts",
        "--",
        "continuous-fuzz.test.ts",
      ],
      { stdout: "pipe" },
    );
    const output = await new Response(process.stdout).text();
    expect(await process.exited).toBe(0);
    expect(JSON.parse(output).files).toEqual(["continuous-fuzz.test.ts"]);
    expect(parseConsumerArgs(["continuous-fuzz.test.ts"], "/repo").files).toEqual([
      "continuous-fuzz.test.ts",
    ]);
    expect(parseConsumerArgs(["--", "continuous-fuzz.test.ts"], "/repo").files).toEqual([
      "continuous-fuzz.test.ts",
    ]);
    expect(() => parseConsumerArgs(["--unknown"], "/repo")).toThrow();
  });

  test("fast-check 4's actual errorInstance survives reporting", () => {
    const details = fc.check(
      fc.property(fc.constant(1), () => {
        throw new Error("sentinel oracle failure");
      }),
      propertyConfig({ numRuns: 1 }),
    );
    expect(classifyFuzzRun(details)).toMatchObject({ status: "finding", completed: 1 });
    const health = classifyFuzzRun(details);
    expect(health.status === "finding" && health.detail).toContain("sentinel oracle failure");
    expect(checkFuzzHealth({ log: line(health), outcome: "failure" }).status).toBe("finding");
    if (!details.failed || details.counterexample === null) panic("Expected a real counterexample");
    // Shrinking may be interrupted after fast-check has already found a failure.
    const interrupted = { ...details, interrupted: true } satisfies typeof details;
    expect(classifyFuzzRun(interrupted)).toMatchObject({ status: "finding", completed: 1 });
    for (const errorInstance of [undefined, null, "undefined", "null", " "]) {
      const invalid = { ...interrupted, errorInstance } satisfies typeof details;
      expect(classifyFuzzRun(invalid).status).toBe("infrastructure");
    }
  });

  test("zero runs, exhausted skips and interrupted runs are infrastructure", () => {
    const property = fc.property(fc.constant(1), () => true);
    expect(classifyFuzzRun(fc.check(property, propertyConfig({ numRuns: 0 }))).status).toBe(
      "infrastructure",
    );
    expect(
      classifyFuzzRun(
        fc.check(
          fc.property(fc.constant(1), () => {
            fc.pre(false);
          }),
          propertyConfig({ numRuns: 1, maxSkipsPerRun: 1 }),
        ),
      ).status,
    ).toBe("infrastructure");
    expect(
      classifyFuzzRun(
        fc.check(
          property,
          propertyConfig({ numRuns: 1, interruptAfterTimeLimit: 0, markInterruptAsFailure: true }),
        ),
      ).status,
    ).toBe("infrastructure");
  });

  test("crash before a case, timeout, absent and malformed reports raise alerts", () => {
    for (const outcome of ["failure", "cancelled", "skipped", "success"]) {
      expect(checkFuzzHealth({ log: "startup crashed", outcome }).status).toBe("infrastructure");
    }
    const passed = line({ status: "passed", completed: 4 });
    expect(checkFuzzHealth({ log: passed, outcome: "success" }).status).toBe("passed");
    expect(checkFuzzHealth({ log: passed, outcome: "failure" }).status).toBe("infrastructure");
    for (const report of [
      { status: "passed", completed: 0 },
      { status: "finding", completed: 1, detail: "undefined" },
      { status: "finding", completed: 1, detail: " null " },
      { status: "finding", completed: 1, detail: "" },
      { status: "wat", completed: 1, detail: "wat" },
    ]) {
      expect(checkFuzzHealth({ log: line(report), outcome: "failure" }).status).toBe(
        "infrastructure",
      );
    }
    expect(
      checkFuzzHealth({ log: `${passed}\n${HEALTH_MARKER}{`, outcome: "success" }).status,
    ).toBe("infrastructure");
    expect(
      checkFuzzHealth({
        log: `${line({ status: "finding", completed: 1, detail: "oracle" })}\n${line({ status: "started", completed: 0 })}`,
        outcome: "failure",
      }).status,
    ).toBe("infrastructure");
  });
});
