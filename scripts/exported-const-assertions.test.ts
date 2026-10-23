import { describe, expect, setDefaultTimeout, test } from "bun:test";

import { scanTestTimeout } from "../test/scan-timeout";
import { exportedConstAssertionsWithoutTypes } from "./lib/exported-const-assertions";

setDefaultTimeout(scanTestTimeout(30_000));

describe("exported const assertion guard", () => {
  test("published package sources satisfy the ratchet", async () => {
    const result = await Bun.$`bun scripts/check-exported-const-assertions.ts`
      .quiet()
      .nothrow();
    expect(result.exitCode).toBe(0);
  });

  test("rejects the source-carrier declaration without its explicit type", () => {
    const findings = exportedConstAssertionsWithoutTypes(
      'export const SOURCE_CARRIER_MARK_NAMES = [RUN_IDENTITY_MARK_NAME, "hyperlink"] as const;',
    );
    expect(findings.map(({ name }) => name)).toEqual([
      "SOURCE_CARRIER_MARK_NAMES",
    ]);
  });

  test("accepts explicitly typed arrays and objects", () => {
    expect(
      exportedConstAssertionsWithoutTypes(
        'export const NAMES: readonly Name[] = ["first", "second"] as const;\n' +
          "export const LOOKUP: Readonly<Record<Name, number>> = { first: 1, second: 2 } as const;",
      ),
    ).toEqual([]);
  });
});
