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

type ExpiringIgnore = {
  advisory: string;
  packageName: string;
  /** Last vulnerable version; a newer published version is a fix. */
  vulnerableThrough: string;
  /** Latest-release check of the package that pulls the vulnerable one in. */
  parent: string;
  /** ISO date after which the ignore no longer applies. */
  expires: string;
};

/** Ignores without an expiry predate this mechanism. */
const PERMANENT_IGNORES = ["GHSA-5p4m-2wfm-xmqj"];

const EXPIRING_IGNORES: readonly ExpiringIgnore[] = [
  {
    // 2026-10-02: node-forge <=1.4.0 accepts malformed RSA PKCS#1 v1.5
    // signatures. It reaches the tree only through development tooling
    // (@stll/folio-nuxt devDependencies nuxt and @nuxt/module-builder ->
    // @nuxt/cli -> listhen, which uses it for local HTTPS certificates). No
    // published package depends on it, and folio never verifies RSA
    // signatures. No patched node-forge exists yet. Tracked in #1344.
    advisory: "GHSA-86w9-cpqp-85rv",
    packageName: "node-forge",
    vulnerableThrough: "1.4.0",
    parent: "listhen",
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

/** Why an expiring ignore no longer applies, or undefined while it does. */
export const staleIgnoreReason = (
  ignore: ExpiringIgnore,
  state: RegistryState,
  today: string,
): string | undefined => {
  if (today > ignore.expires) {
    return `${ignore.advisory}: the ignore expired on ${ignore.expires}; re-check the advisory and update or renew it.`;
  }
  if (compareVersions(state.latestVersion, ignore.vulnerableThrough) > 0) {
    return `${ignore.advisory}: ${ignore.packageName}@${state.latestVersion} is published; update it and remove the ignore.`;
  }
  if (!state.stillRequiredByParent) {
    return `${ignore.advisory}: the latest ${ignore.parent} no longer requires ${ignore.packageName}; update it and remove the ignore.`;
  }
  return undefined;
};

type Manifest = { version: string; dependencies?: Record<string, string> };

const latestManifest = async (name: string): Promise<Manifest> => {
  const response = await fetch(`https://registry.npmjs.org/${name}/latest`);
  if (!response.ok) {
    throw new Error(`npm registry returned ${response.status} for ${name}`);
  }
  return (await response.json()) as Manifest;
};

const main = async (): Promise<number> => {
  const today = new Date().toISOString().slice(0, 10);
  const stale: string[] = [];
  for (const ignore of EXPIRING_IGNORES) {
    const [vulnerable, parent] = await Promise.all([
      latestManifest(ignore.packageName),
      latestManifest(ignore.parent),
    ]);
    const reason = staleIgnoreReason(
      ignore,
      {
        latestVersion: vulnerable.version,
        stillRequiredByParent: ignore.packageName in (parent.dependencies ?? {}),
      },
      today,
    );
    if (reason) stale.push(reason);
  }
  if (stale.length > 0) {
    for (const reason of stale) console.error(reason);
    return 1;
  }
  const ignores = [...PERMANENT_IGNORES, ...EXPIRING_IGNORES.map((ignore) => ignore.advisory)];
  const audit = Bun.spawn(
    ["bun", "audit", "--audit-level=moderate", ...ignores.map((id) => `--ignore=${id}`)],
    { stdout: "inherit", stderr: "inherit" },
  );
  return await audit.exited;
};

if (import.meta.main) {
  process.exit(await main());
}
