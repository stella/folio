/**
 * Dependency audit with a short, self-expiring ignore list.
 *
 * An advisory is ignored only while it has no fix we can take. Each entry
 * names the path it arrives by and why it cannot reach folio's runtime, and
 * the audit fails again as soon as a fix is published or the entry expires,
 * so an ignore can never outlive its reason.
 */

type RegistryState = {
  /** Latest published version of the advisory's package. */
  latestVersion: string;
  /** Whether the latest release of the dependency that pulls it in still requires it. */
  stillRequiredByParent: boolean;
};

/**
 * Why the vulnerable package stays in the tree, checked against the latest
 * release of `parent`. `through` covers a parent that pins an intermediate
 * dependency to a frozen release line which still requires the package.
 */
type Requirement =
  | { type: "direct"; parent: string }
  | { type: "through"; parent: string; dependency: string; range: string };

type ExpiringIgnore = {
  advisory: string;
  packageName: string;
  /** Last vulnerable version; a newer published version is a fix. */
  vulnerableThrough: string;
  requiredBy: Requirement;
  /** ISO date after which the ignore no longer applies. */
  expires: string;
};

/** Ignores without an expiry predate this mechanism. */
const PERMANENT_IGNORES = ["GHSA-5p4m-2wfm-xmqj"];

const EXPIRING_IGNORES: readonly ExpiringIgnore[] = [
  {
    // 2026-10-03: braces <=3.0.3 exhausts the stack on deeply nested
    // patterns. It reaches the tree only through development tooling
    // (@tailwindcss/cli, nuxt and @nuxt/module-builder -> @parcel/watcher ->
    // micromatch). No published package depends on it, and folio never
    // expands untrusted patterns. No patched braces exists yet.
    advisory: "GHSA-vfj7-8cjw-p6xm",
    packageName: "braces",
    vulnerableThrough: "3.0.3",
    requiredBy: { type: "direct", parent: "micromatch" },
    expires: "2026-11-01",
  },
  {
    // 2026-10-06: no patched sprintf-js is published. It arrives through
    // @rushstack/ts-command-line's argparse ~1.0.9 pin.
    advisory: "GHSA-hp3w-g68c-fv3c",
    packageName: "sprintf-js",
    vulnerableThrough: "1.1.3",
    requiredBy: {
      type: "through",
      parent: "@rushstack/ts-command-line",
      dependency: "argparse",
      range: "~1.0.9",
    },
    expires: "2026-11-01",
  },
];

const compareVersions = (left: string, right: string): number => {
  const parse = (version: string) =>
    version
      .split("-")[0]!
      .split(".")
      .map((part) => Number.parseInt(part, 10) || 0);
  const a = parse(left);
  const b = parse(right);
  for (let index = 0; index < Math.max(a.length, b.length); index += 1) {
    const difference = (a[index] ?? 0) - (b[index] ?? 0);
    if (difference !== 0) return difference;
  }
  return 0;
};

/**
 * Why an expiring ignore no longer applies, or undefined while it does.
 * Without a registry `state` (an outage), only the expiry date is checked.
 */
export const staleIgnoreReason = (
  ignore: ExpiringIgnore,
  state: RegistryState | undefined,
  today: string,
): string | undefined => {
  if (today > ignore.expires) {
    return `${ignore.advisory}: the ignore expired on ${ignore.expires}; re-check the advisory and update or renew it.`;
  }
  if (state === undefined) return undefined;
  if (compareVersions(state.latestVersion, ignore.vulnerableThrough) > 0) {
    return `${ignore.advisory}: ${ignore.packageName}@${state.latestVersion} is published; update it and remove the ignore.`;
  }
  if (!state.stillRequiredByParent) {
    return `${ignore.advisory}: the latest ${ignore.requiredBy.parent} no longer requires ${ignore.packageName}; update it and remove the ignore.`;
  }
  return undefined;
};

type Manifest = { version: string; dependencies?: Record<string, string> };

/** Whether the parent's latest manifest still pulls the vulnerable package in. */
export const isStillRequired = (
  ignore: ExpiringIgnore,
  parent: Manifest,
): boolean => {
  const dependencies = parent.dependencies ?? {};
  const { requiredBy } = ignore;
  switch (requiredBy.type) {
    case "direct":
      return ignore.packageName in dependencies;
    case "through":
      return dependencies[requiredBy.dependency] === requiredBy.range;
    default: {
      const exhaustive: never = requiredBy;
      return exhaustive;
    }
  }
};

/** A registry answer, or an outage that outlasted every retry. */
export type ManifestLookup =
  | { type: "manifest"; manifest: Manifest }
  | { type: "unavailable"; reason: string };

/**
 * The latest manifest of `name`. Network errors, 429 and 5xx answers are
 * retried after each of `delaysMs`; an outage that outlasts them is reported
 * as unavailable. Any other answer is an error: the name or URL is wrong.
 */
export const fetchLatestManifest = async (
  name: string,
  fetcher: (url: string) => Promise<Response> = fetch,
  delaysMs: readonly number[] = [1_000, 3_000],
): Promise<ManifestLookup> => {
  let reason = "";
  for (let attempt = 0; attempt <= delaysMs.length; attempt++) {
    if (attempt > 0) await Bun.sleep(delaysMs[attempt - 1] ?? 0);
    let response: Response;
    try {
      response = await fetcher(`https://registry.npmjs.org/${name}/latest`);
    } catch (error) {
      reason = `request failed: ${String(error)}`;
      continue;
    }
    if (response.ok)
      return {
        type: "manifest",
        manifest: (await response.json()) as Manifest,
      };
    if (response.status !== 429 && response.status < 500) {
      throw new Error(`npm registry returned ${response.status} for ${name}`);
    }
    reason = `npm registry returned ${response.status}`;
  }
  return { type: "unavailable", reason };
};

const main = async (): Promise<number> => {
  const today = new Date().toISOString().slice(0, 10);
  const stale: string[] = [];
  for (const ignore of EXPIRING_IGNORES) {
    const [vulnerable, parent] = await Promise.all([
      fetchLatestManifest(ignore.packageName),
      fetchLatestManifest(ignore.requiredBy.parent),
    ]);
    let state: RegistryState | undefined;
    if (vulnerable.type === "manifest" && parent.type === "manifest") {
      state = {
        latestVersion: vulnerable.manifest.version,
        stillRequiredByParent: isStillRequired(ignore, parent.manifest),
      };
    } else {
      const outage = vulnerable.type === "unavailable" ? vulnerable : parent;
      if (outage.type === "unavailable") {
        console.warn(
          `${ignore.advisory}: could not check for a published fix (${outage.reason}); the ignore stays until ${ignore.expires}.`,
        );
      }
    }
    const reason = staleIgnoreReason(ignore, state, today);
    if (reason) stale.push(reason);
  }
  if (stale.length > 0) {
    for (const reason of stale) console.error(reason);
    return 1;
  }
  const ignores = [
    ...PERMANENT_IGNORES,
    ...EXPIRING_IGNORES.map((ignore) => ignore.advisory),
  ];
  const audit = Bun.spawn(
    [
      "bun",
      "audit",
      "--audit-level=moderate",
      ...ignores.map((id) => `--ignore=${id}`),
    ],
    { stdout: "inherit", stderr: "inherit" },
  );
  return await audit.exited;
};

if (import.meta.main) {
  process.exit(await main());
}
