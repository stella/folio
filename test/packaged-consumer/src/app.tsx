// Runtime half of the packaged-consumer gate. Where `main.ts` proves a
// production `vite build` RESOLVES everything the packed @stll/folio-* tarballs
// reference, this file proves the built output RUNS: it mounts the packaged
// `DocxEditor` with the shipped `@stll/folio-react/messages` catalog and shipped
// `editor.css`, loads a real .docx fixture, and exposes a browser-driven test
// hook (`window.__folioSmoke`). The Playwright smoke (`smoke.spec.ts`) drives it
// and fails on any console error — catching the two runtime-only failure classes
// the build gate cannot see: a font-metrics worker that resolves at build time
// but fails to spawn/execute, and UI strings missing from the runtime catalog
// (`IntlError: MISSING_MESSAGE`).
//
// Everything here is imported from the INSTALLED tarballs (`@stll/folio-*`), not
// repo source, so this exercises the published artifacts a downstream app gets.

import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { DocxEditor } from "@stll/folio-react";
import type { DocxEditorRef } from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";
import {
  canPrefetchMeasurement,
  prefetchMeasurement,
} from "@stll/folio-core/layout-engine/measure/measureWorker";
import { getCachedTextWidth } from "@stll/folio-core/layout-engine/measure/cache";
import { WORKER_FONT_FINGERPRINT_TEXT } from "@stll/folio-core/layout-engine/measure/measureWorkerProtocol";

import "@stll/folio-react/editor.css";

import { useEffect, useRef, useState } from "react";

const LOCALE = "en";

// Turn ON the off-main-thread font-metrics worker. It is OFF by default and is a
// best-effort cache pre-warm (never on the critical layout path), so mounting
// the editor alone does not construct the worker. Enabling the flag makes the
// editor's real layout path call `prefetchMeasurement`, which constructs the
// SHIPPED worker via `new Worker(new URL("font-metrics.worker.js", ...))` inside
// the packed dist — the exact code path the smoke must exercise at runtime.
globalThis.__folioFeatureFlags = { workerFontMetrics: true };

// Drive one real measurement round-trip through the shipped worker and return
// the width it computed. This is the genuinely worker-load-bearing probe: the
// value only lands in the main-thread cache if the worker actually SPAWNED,
// received the batch, ran measurement, and posted a reply back. A worker that
// fails to spawn/execute (bad URL, module throws on load) fires the proxy's
// `error` handler, marks the proxy dead, and this never resolves a width.
//
// `sans-serif` (not a bundled web font) is measured identically by the
// main-thread 2D canvas and the worker's OffscreenCanvas, so the worker's
// font-fingerprint guard passes and the entry is not skipped. A unique probe
// text + cache key guarantees the read reflects THIS round-trip, not a warm
// cache from the editor's own layout.
async function measureRoundTrip(): Promise<{ width: number; alive: boolean }> {
  const font = "16px sans-serif";
  const letterSpacing = 0;
  const nonce = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const fontCacheKey = `folio-smoke|${font}|${nonce}`;
  const text = `folio-smoke-probe-${nonce}`;

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (ctx === null) {
    return { width: -1, alive: canPrefetchMeasurement() };
  }
  ctx.font = font;
  const fontFingerprintWidth = ctx.measureText(WORKER_FONT_FINGERPRINT_TEXT).width;

  prefetchMeasurement(text, font, letterSpacing, 1, fontCacheKey, fontFingerprintWidth);

  const deadline = Date.now() + 8000;
  while (Date.now() < deadline) {
    const width = getCachedTextWidth(text, fontCacheKey, letterSpacing);
    if (width !== undefined) {
      return { width, alive: canPrefetchMeasurement() };
    }
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  return { width: -1, alive: canPrefetchMeasurement() };
}

declare global {
  var __folioFeatureFlags: { workerFontMetrics?: boolean } | undefined;
  var __folioSmoke:
    | {
        getEditorRef: () => DocxEditorRef | null;
        measureRoundTrip: () => Promise<{ width: number; alive: boolean }>;
      }
    | undefined;
}

function App() {
  const editorRef = useRef<DocxEditorRef>(null);
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);

  useEffect(() => {
    void (async () => {
      const response = await fetch("/fixture.docx");
      setBuffer(await response.arrayBuffer());
    })();
  }, []);

  useEffect(() => {
    globalThis.__folioSmoke = {
      getEditorRef: () => editorRef.current,
      measureRoundTrip,
    };
    return () => {
      globalThis.__folioSmoke = undefined;
    };
  }, []);

  return (
    <IntlProvider
      locale={LOCALE}
      messages={getFolioMessages(LOCALE)}
      timeZone={Intl.DateTimeFormat().resolvedOptions().timeZone}
    >
      <DocxEditor
        ref={editorRef}
        documentBuffer={buffer}
        author="Folio Smoke"
        showToolbar={true}
        initialZoom={1}
      />
    </IntlProvider>
  );
}

const container = document.querySelector("#app");
if (container) {
  createRoot(container).render(<App />);
}
