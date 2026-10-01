import { expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { viewportPosition } from "./viewportPosition";

test(
  "viewport overlays reconstruct painted client coordinates across zoom and host scroll",
  () => {
    fc.assert(
      fc.property(
        fc.record({
          left: fc.integer({ min: -10000, max: 10000 }),
          top: fc.integer({ min: -10000, max: 10000 }),
          x: fc.integer({ min: 0, max: 5000 }),
          y: fc.integer({ min: 0, max: 5000 }),
          zoomPercent: fc.integer({ min: 25, max: 400 }),
        }),
        ({ left, top, x, y, zoomPercent }) => {
          const zoom = zoomPercent / 100;
          const clientX = left + x * zoom;
          const clientY = top + y * zoom;
          const local = viewportPosition({ clientX, clientY, viewport: { left, top } });
          // Positioning a sibling in the unscaled viewport must reconstruct the
          // painted point, regardless of page scale or the host's scroll origin.
          expect(left + local.left).toBeCloseTo(clientX, 8);
          expect(top + local.top).toBeCloseTo(clientY, 8);
          expect(local.left).toBeCloseTo(x * zoom, 8);
          expect(local.top).toBeCloseTo(y * zoom, 8);
        },
      ),
      propertyConfig({ seed: 1330, numRuns: 100 }),
    );
  },
  propertyTestTimeout(10_000),
);
