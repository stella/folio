import path from "node:path";

const targets = (value: unknown): string[] => {
  if (typeof value === "string") return [value];
  if (value === null) return [];
  if (Array.isArray(value)) return value.flatMap(targets);
  if (typeof value !== "object")
    throw new TypeError("Invalid package export target");
  return Object.values(value).flatMap(targets);
};

const patternMatch = (pattern: string, value: string) => {
  const star = pattern.indexOf("*");
  if (star < 0) return pattern === value ? "" : null;
  const prefix = pattern.slice(0, star);
  const suffix = pattern.slice(star + 1);
  return value.startsWith(prefix) &&
    value.endsWith(suffix) &&
    value.length >= prefix.length + suffix.length
    ? value.slice(prefix.length, value.length - suffix.length)
    : null;
};

const hasCondition = (value: unknown, condition: string): boolean => {
  if (typeof value === "string") {
    if (condition === "types") return /\.d\.[cm]?ts$/u.test(value);
    return (
      condition === "import" ||
      (condition === "require" && value.endsWith(".cjs"))
    );
  }
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value))
    return value.some((entry) => hasCondition(entry, condition));
  return Object.entries(value).some(([key, entry]) => {
    if (key === condition) return targets(entry).length > 0;
    if (
      key === "default" &&
      (condition === "import" || condition === "require")
    )
      return targets(entry).length > 0;
    // Loader-specific conditions cannot activate another loader or type-only branch.
    if (["types", "import", "require", "default"].includes(key)) return false;
    return hasCondition(entry, condition);
  });
};

type PublicEntriesOptions = {
  name: string;
  exports: Record<string, unknown>;
  files: readonly string[];
};

/** Expand wildcard exports against packed files, respecting exact and null overrides. */
export const publicEntries = ({
  name,
  exports,
  files,
}: PublicEntriesOptions) => {
  const candidates = new Set<string>();
  for (const [subpath, value] of Object.entries(exports)) {
    if (value === null) continue;
    if (!subpath.includes("*")) {
      candidates.add(subpath);
      continue;
    }
    for (const target of targets(value)) {
      for (const file of files) {
        const capture = patternMatch(target, `./${file}`);
        if (capture !== null) candidates.add(subpath.replaceAll("*", capture));
      }
    }
  }
  return [...candidates].toSorted().flatMap((subpath) => {
    const key = Object.keys(exports)
      .filter((candidate) => patternMatch(candidate, subpath) !== null)
      .toSorted((left, right) => {
        if (left === subpath) return -1;
        if (right === subpath) return 1;
        const prefixDifference = right.indexOf("*") - left.indexOf("*");
        return prefixDifference || right.length - left.length;
      })
      .at(0);
    if (!key) throw new TypeError(`Export has no matching key: ${subpath}`);
    const value = exports[key];
    if (value === null) return [];
    const asset = targets(value).every((target) =>
      [".css", ".wasm", ".json"].includes(path.extname(target)),
    );
    return [
      {
        specifier: subpath === "." ? name : `${name}/${subpath.slice(2)}`,
        esm: !asset && hasCondition(value, "import"),
        cjs: !asset && hasCondition(value, "require"),
        types: !asset && hasCondition(value, "types"),
        asset,
      },
    ];
  });
};
