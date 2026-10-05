import { expect, test } from "bun:test";
import fc from "fast-check";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { assertProperty } from "../../../../test/property-testing";
import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { createEmptyDocument } from "../utils/createDocument";

test("canonical drivers release owned DOM across overlapping mount and teardown sequences", () => {
  // Driver-only tests missed global registration leaking into later test modules.
  const ambient = typeof document !== "undefined";
  const registered = GlobalRegistrator.isRegistered;
  const fetchBefore = globalThis.fetch;
  const urlBefore = globalThis.URL;
  assertProperty(
    fc.property(fc.array(fc.nat(5), { minLength: 1, maxLength: 20 }), (sequence) => {
      const drivers: ReturnType<typeof createCanonicalEditorHarness>[] = [];
      try {
        for (const action of sequence) {
          if (action < 3 || drivers.length === 0) {
            const source = createEmptyDocument({ initialText: "Lifecycle" });
            const paragraph = source.package.document.content.at(0);
            if (paragraph?.type !== "paragraph")
              throw new TypeError("Lifecycle fixture lacks paragraph.");
            paragraph.paraId = "12345678";
            drivers.push(createCanonicalEditorHarness(source, "editing"));
          } else {
            const driver = drivers.splice(action % drivers.length, 1).at(0);
            driver?.dispose();
            driver?.dispose();
          }
          if (drivers.length > 0) expect(typeof document).toBe("object");
        }
      } finally {
        for (const driver of drivers) driver.dispose();
      }
      expect(typeof document !== "undefined").toBe(ambient);
      expect(GlobalRegistrator.isRegistered).toBe(registered);
      expect(globalThis.fetch).toBe(fetchBefore);
      expect(globalThis.URL).toBe(urlBefore);
    }),
    { numRuns: 20 },
  );
});
