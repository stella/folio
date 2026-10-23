import { describe, expect, test } from "bun:test";

import { columnBalancingForCompatibility } from "./noColumnBalanceCompatibility";

describe("w:noColumnBalance compatibility", () => {
  test("honors an enabled option in legacy and undeclared modes", () => {
    expect(columnBalancingForCompatibility(true, undefined)).toBe("natural");
    expect(columnBalancingForCompatibility(true, 14)).toBe("natural");
  });

  test("ignores the option in mode 15 and when absent or disabled", () => {
    expect(columnBalancingForCompatibility(true, 15)).toBe("balanced");
    expect(columnBalancingForCompatibility(false, 14)).toBe("balanced");
    expect(columnBalancingForCompatibility(undefined, 14)).toBe("balanced");
  });
});
