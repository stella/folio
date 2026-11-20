/**
 * Browser tests read the playground ports from tests/parity/playgroundHosts.ts
 * alone: it takes free ports when none are set and shares them with the
 * Playwright workers and servers. A test that reads the port variables or
 * names a fixed port itself would drift from the servers it talks to.
 */
import { expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OWNER = "tests/parity/playgroundHosts.ts";
const PORT_READ =
  /FOLIO_PLAYGROUND_PORT|FOLIO_PLAYGROUND_VUE_PORT|localhost:420[01]\b/u;

test("only playgroundHosts.ts reads or names the playground ports", () => {
  const offenders = readdirSync(path.join(REPO_ROOT, "tests"), {
    recursive: true,
  })
    .map((entry) =>
      path.posix.join("tests", String(entry).replaceAll("\\", "/")),
    )
    .filter((file) => /\.tsx?$/u.test(file) && file !== OWNER)
    .filter((file) =>
      PORT_READ.test(readFileSync(path.join(REPO_ROOT, file), "utf8")),
    );
  expect(offenders).toEqual([]);
});
