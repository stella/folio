import { describe, expect, test } from "bun:test";
import { readdirSync } from "node:fs";

import { installParityMeasurement, layoutBothPaths } from "./__tests__/headlessLayoutParity";

installParityMeasurement();

const fixtureRoots = [
  { label: "visual", url: new URL("../../../../tests/visual/fixtures/", import.meta.url) },
  { label: "corpus", url: new URL("../docx/__tests__/__fixtures__/corpus/", import.meta.url) },
];

// There are no fixture exclusions: a load failure in either path fails its
// named case. Any future unsupported fixture must have an explicit reason.
describe("headless/editor fixture page-break parity", () => {
  for (const { label, url } of fixtureRoots) {
    const names = readdirSync(url)
      .filter((name) => name.endsWith(".docx"))
      .sort();
    test(`fixture inventory is nonempty: ${label}`, () => {
      expect(names.length).toBeGreaterThan(0);
    });
    for (const name of names) {
      const fixture = new URL(name, url);
      test(`${label}/${name}`, async () => {
        await layoutBothPaths(await Bun.file(fixture).arrayBuffer());
      }, 60_000);
    }
  }
});
