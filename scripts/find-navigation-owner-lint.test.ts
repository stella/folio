import { describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { z } from "zod";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_ID = "folio-find-navigation(navigation-owner)";
const IMPORT = 'import { createFindMatchView } from "@stll/folio-core/utils/findMatchView";';
const CASES = {
  vue: {
    source: `${IMPORT}
    let matchView: ReturnType<typeof createFindMatchView> | null = null;
    export const go = () => matchView?.reveal(0);
    matchView = createFindMatchView({});`,
    count: 1,
  },
  react: {
    source: `${IMPORT}
    import { useRef } from "react";
    const matchView = useRef<ReturnType<typeof createFindMatchView> | null>(null);
    export const go = () => matchView.current?.reveal(0);`,
    count: 1,
  },
  aliases: {
    source:
      'import { createFindMatchView as create } from "@stll/folio-core/utils/findMatchView";\n' +
      `
    const view = create({});
    const alias = view;
    export const go = () => alias["reveal"](0);
    export const { reveal: raw } = alias;`,
    count: 2,
  },
  assigned: {
    source: `${IMPORT}
    let view = null;
    export const go = () => view?.reveal(0);
    view = true ? createFindMatchView({}) : null;`,
    count: 1,
  },
  fallback: {
    source: `${IMPORT}
    import { scrollToMatch as scroll } from "./findReplaceUtils";
    export const view = createFindMatchView({});
    export const go = () => scroll(null, {});`,
    count: 1,
  },
  namespace: {
    source: `import * as search from "@stll/folio-core/utils/findMatchView";
      export const go = () => search.createFindMatchView({}).reveal(0);`,
    count: 1,
  },
  pageSelection: {
    source: `${IMPORT}
      export const view = createFindMatchView({});
      export const go = (root, page) => root.querySelector(\`[data-page-number="\${page}"]\`);`,
    count: 1,
  },
  navigation: {
    source: `${IMPORT}
    const view = createFindMatchView({});
    export const go = () => view.navigate(0);
    export const paint = () => view.paint(0);`,
    count: 0,
  },
  unrelated: {
    source: `${IMPORT}
    export const view = createFindMatchView({});
    const other = { reveal: () => true };
    export const go = () => other.reveal();`,
    count: 0,
  },
};

const lintFixture = (source: string): string[] => {
  const directory = mkdtempSync(path.join(REPO_ROOT, "test", "__fixtures__", "find-navigation-"));
  try {
    const fixture = path.join(directory, "fixture.ts");
    writeFileSync(fixture, source);
    const result = Bun.spawnSync(
      [
        "bun",
        "--bun",
        "oxlint",
        "-c",
        "oxlint.fixtures.config.ts",
        "--no-ignore",
        "--format",
        "json",
        fixture,
      ],
      { cwd: REPO_ROOT },
    );
    const { diagnostics } = z
      .object({ diagnostics: z.array(z.object({ code: z.string() })) })
      .parse(JSON.parse(result.stdout.toString()));
    return diagnostics.map(({ code }) => code).filter((code) => code === RULE_ID);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

describe("find navigation ownership", () => {
  for (const [name, { source, count }] of Object.entries(CASES))
    test(
      name,
      () => expect(lintFixture(source)).toEqual(Array.from({ length: count }, () => RULE_ID)),
      30_000,
    );
});
