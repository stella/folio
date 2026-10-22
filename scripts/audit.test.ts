import { describe, expect, test } from "bun:test";

import { fetchLatestManifest, isStillRequired, staleIgnoreReason } from "./audit";

const ignore = {
  advisory: "GHSA-test",
  packageName: "vulnerable",
  vulnerableThrough: "1.4.0",
  requiredBy: { type: "direct", parent: "parent" },
  expires: "2026-11-01",
} as const;

const throughIgnore = {
  ...ignore,
  requiredBy: { type: "through", parent: "parent", dependency: "middle", range: "~1.0.9" },
} as const;

describe("self-expiring audit ignores", () => {
  test("an ignore applies while no fix is published and before it expires", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toBeUndefined();
  });

  test("a published fix ends the ignore", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.1", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toContain("vulnerable@1.4.1 is published");
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "2.0.0", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toContain("is published");
  });

  test("a parent release that drops the package ends the ignore", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: false },
        "2026-10-02",
      ),
    ).toContain("no longer requires vulnerable");
  });

  test("the ignore expires on its date", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-11-01",
      ),
    ).toBeUndefined();
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-11-02",
      ),
    ).toContain("expired on 2026-11-01");
  });

  test("a direct parent requires the package while it declares it", () => {
    const version = "1.0.0";
    expect(isStillRequired(ignore, { version, dependencies: { vulnerable: "^1.0.0" } })).toBe(true);
    expect(isStillRequired(ignore, { version, dependencies: { other: "^1.0.0" } })).toBe(false);
    expect(isStillRequired(ignore, { version })).toBe(false);
  });

  test("a transitive requirement holds only while the parent keeps the pinned range", () => {
    const version = "1.0.0";
    expect(isStillRequired(throughIgnore, { version, dependencies: { middle: "~1.0.9" } })).toBe(
      true,
    );
    expect(isStillRequired(throughIgnore, { version, dependencies: { middle: "^2.0.0" } })).toBe(
      false,
    );
    expect(isStillRequired(throughIgnore, { version, dependencies: {} })).toBe(false);
  });
});

describe("registry lookups for expiring ignores", () => {
  const answer = (status: number, body: unknown = { version: "1.5.0" }) =>
    new Response(JSON.stringify(body), { status });
  const scripted = (responses: (Response | Error)[]) => {
    const calls: string[] = [];
    const fetcher = (url: string): Promise<Response> => {
      calls.push(url);
      const next = responses.shift();
      if (next === undefined) throw new Error("no scripted response left");
      return next instanceof Error ? Promise.reject(next) : Promise.resolve(next);
    };
    return { calls, fetcher };
  };

  test("an outage keeps the ignore until its expiry date, which still applies", () => {
    expect(staleIgnoreReason(ignore, undefined, "2026-10-22")).toBeUndefined();
    expect(staleIgnoreReason(ignore, undefined, "2026-11-02")).toContain("expired on 2026-11-01");
  });

  test("a transient 503 is retried and the next answer is used", async () => {
    const { calls, fetcher } = scripted([answer(503), answer(200)]);
    expect(await fetchLatestManifest("vulnerable", fetcher, [0, 0])).toEqual({
      type: "manifest",
      manifest: { version: "1.5.0" },
    });
    expect(calls).toHaveLength(2);
  });

  test("an outage that outlasts every retry is reported as unavailable", async () => {
    const { calls, fetcher } = scripted([answer(503), new Error("socket hang up"), answer(502)]);
    expect(await fetchLatestManifest("vulnerable", fetcher, [0, 0])).toEqual({
      type: "unavailable",
      reason: "npm registry returned 502",
    });
    expect(calls).toHaveLength(3);
  });

  test("a client error is not an outage", async () => {
    const { fetcher } = scripted([answer(404)]);
    await expect(fetchLatestManifest("missing", fetcher, [0, 0])).rejects.toThrow(
      "returned 404 for missing",
    );
  });
});
