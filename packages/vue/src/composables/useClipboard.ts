/**
 * Vue port of packages/react/src/hooks/useClipboard.ts. Wraps the
 * framework-agnostic clipboard helpers in core.
 *
 * PORT-BLOCKED (copy/cut + DOM selection extraction): upstream's copy/cut
 * surface depends on `@stll/folio-core` APIs that are absent from our fork:
 *   - `managers/ClipboardManager` — `ClipboardSelection`, `getSelectionRuns`,
 *     `createSelectionFromDOM` (our core has no ClipboardManager, so there is
 *     no DOM-run-extraction primitive to build `copy`/`cut` on),
 *   - a `theme` field on `copyRuns`'s options (our `ClipboardOptions` only
 *     exposes `includeFormatting`/`cleanWordFormatting`/`onError`), and
 *   - a `Theme` value/type export from the core root.
 * `copyRuns`/`runsToClipboardContent` themselves DO exist, but the selection
 * capture they consume does not, so the copy/cut half is omitted rather than
 * fabricated. Only the paste path (which uses the available `parseClipboardHtml`)
 * is ported below. Restore `copy`/`cut` once a DOM-run-extraction API lands in
 * core.
 */
import { ref, type Ref } from "vue";
import {
  parseClipboardHtml,
  type ParsedClipboardContent,
} from "@stll/folio-core/utils/clipboard";

export interface UseClipboardOptions {
  onPaste?: (content: ParsedClipboardContent, asPlainText: boolean) => void;
  cleanWordFormatting?: boolean;
  editable?: boolean;
  onError?: (error: Error) => void;
}

export interface UseClipboardReturn {
  paste: (asPlainText?: boolean) => Promise<ParsedClipboardContent | null>;
  isProcessing: Ref<boolean>;
  lastPastedContent: Ref<ParsedClipboardContent | null>;
}

export function useClipboard(options: UseClipboardOptions = {}): UseClipboardReturn {
  const { onPaste, cleanWordFormatting = true, editable = true, onError } = options;

  const isProcessing = ref(false);
  const lastPastedContent = ref<ParsedClipboardContent | null>(null);

  async function paste(asPlainText = false): Promise<ParsedClipboardContent | null> {
    if (isProcessing.value || !editable) return null;
    isProcessing.value = true;
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        let html = "";
        let plainText = "";
        for (const item of items) {
          if (item.types.includes("text/html")) {
            html = await (await item.getType("text/html")).text();
          }
          if (item.types.includes("text/plain")) {
            plainText = await (await item.getType("text/plain")).text();
          }
        }
        if (asPlainText) html = "";
        const content = parseClipboardHtml(html, plainText, cleanWordFormatting);
        lastPastedContent.value = content;
        onPaste?.(content, asPlainText);
        return content;
      }
      return null;
    } catch (e) {
      onError?.(e instanceof Error ? e : new Error(String(e)));
      return null;
    } finally {
      isProcessing.value = false;
    }
  }

  return {
    paste,
    isProcessing,
    lastPastedContent,
  };
}
