import { describe, expect, test } from "bun:test";

import {
  describeException,
  expiringWithin,
  quarantineExceptions,
  toolchainExceptions,
} from "./expiring-exceptions";

const NOW = new Date("2026-11-14T12:00:00.000Z");

describe("dated exceptions", () => {
  test("every toolchain entry with an expiry is found, whatever its section", () => {
    const found = toolchainExceptions({
      rustExceptions: [
        { pin: "1.97.1", reason: "r", tracking: "t", expires: "2026-11-18" },
      ],
      rustInstallFixtures: [
        {
          path: "scripts/install.sh",
          pin: "1.98.0",
          reason: "r",
          tracking: "t",
          expires: "2027-01-01",
        },
      ],
      future: { nested: [{ expires: "2026-11-15" }] },
      typescriptAllowlist: [{ name: "typescript", spec: "6.0.3" }],
    });
    expect(found.map(({ name, expires }) => [name, expires])).toEqual([
      ["rustExceptions[0] 1.97.1", "2026-11-18"],
      ["rustInstallFixtures[0] scripts/install.sh", "2027-01-01"],
      ["future.nested[0]", "2026-11-15"],
    ]);
    expect(() => toolchainExceptions({ x: [{ expires: "soon" }] })).toThrow(
      "unreadable expiry",
    );
  });

  test("the quarantine excludes carry their bunfig expiry", () => {
    const bunfig = [
      "[install]",
      "minimumReleaseAgeExcludes = [",
      '  "@stll/oxlint-config",',
      '  "left-pad", # quarantine-expires: 2026-11-16T09:00:00.000Z',
      "]",
      "",
    ].join("\n");
    expect(quarantineExceptions(bunfig, "bunfig.toml")).toEqual([
      {
        file: "bunfig.toml",
        name: "left-pad",
        expires: "2026-11-16T09:00:00.000Z",
        expiresMs: Date.parse("2026-11-16T09:00:00.000Z"),
      },
    ]);
  });

  test("only entries inside the window are reported, soonest first, expired included", () => {
    const at = (name: string, expires: string) => ({
      file: "stll-toolchain.json",
      name,
      expires,
      expiresMs: Date.parse(expires),
    });
    const soon = expiringWithin(
      [
        at("later", "2026-11-30"),
        at("inside", "2026-11-20"),
        at("expired", "2026-11-10"),
        at("edge", "2026-11-21T11:59:59.000Z"),
      ],
      NOW,
      7,
    );
    expect(soon.map(({ name }) => name)).toEqual(["expired", "inside", "edge"]);
    expect(describeException(at("inside", "2026-11-20"), NOW)).toEqual({
      annotation:
        "::warning file=stll-toolchain.json,title=Dated exception expires soon::inside expires 2026-11-20 (in 6 d); renew or remove it before the checks fail every run.",
      row: "| stll-toolchain.json | inside | expires 2026-11-20 (in 6 d) |",
    });
    expect(
      describeException(at("expired", "2026-11-10"), NOW).annotation,
    ).toContain("title=Dated exception expired::expired expired 2026-11-10");
  });
});
