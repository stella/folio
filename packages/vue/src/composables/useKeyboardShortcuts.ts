/**
 * Page-level keyboard-shortcut composable — installs a window-level keydown
 * listener that threads through zoom shortcuts (Ctrl+= / Ctrl+- / Ctrl+0)
 * and toggles Find/Replace (Ctrl+F / Ctrl+H) and Hyperlink (Ctrl+K).
 * Ownership of the listener lives here so the SFC stays out of the lifecycle
 * wiring.
 *
 * `scope` decides which presses reach that dispatch: every press on the page,
 * only presses inside the editor, or none at all when the host binds these
 * keys itself. The predicate is shared with the React adapter.
 */

import { onMounted, onBeforeUnmount, type Ref } from "vue";

import { isKeydownInShortcutScope } from "@stll/folio-core/managers/editorShortcuts";
import type { KeyboardShortcutScope } from "@stll/folio-core/managers/editorShortcuts";

export type UseKeyboardShortcutsOptions = {
  showFindReplace: Ref<boolean>;
  showHyperlink: Ref<boolean>;
  /** From useZoom — handles Ctrl+= / Ctrl+- / Ctrl+0. */
  handleZoomKeyDown: (e: KeyboardEvent) => void;
  /**
   * Which presses these shortcuts answer. Read freshly inside the handler so a
   * host change at runtime is honored.
   */
  scope: () => KeyboardShortcutScope;
  /** Elements a press must land inside under the `"editor"` scope. */
  roots: readonly Ref<HTMLElement | null>[];
  /**
   * Host prop accessor — read freshly inside the handler so a host
   * toggle at runtime is honored. (Capturing the prop value at setup
   * time would freeze it.)
   */
  disableFindReplaceShortcuts?: () => boolean | undefined;
  /**
   * Whether `File > Open` / Cmd+O is enabled. Read freshly so a runtime host
   * toggle is honored. When false (or `onOpenDocument` is absent), Cmd+O is
   * left unhandled so the browser default stands.
   */
  showFileOpen?: () => boolean | undefined;
  /** Opens the DOCX picker — wired to the same hidden input as File > Open. */
  onOpenDocument?: () => void;
};

export function useKeyboardShortcuts(opts: UseKeyboardShortcutsOptions) {
  function handleKeyDown(e: KeyboardEvent) {
    const inScope = isKeydownInShortcutScope(e, {
      scope: opts.scope(),
      roots: opts.roots.map((root) => root.value),
    });
    if (!inScope) return;

    // Zoom shortcuts (Ctrl+=/Ctrl+-/Ctrl+0)
    opts.handleZoomKeyDown(e);

    if (!(e.ctrlKey || e.metaKey)) return;
    if (e.key === "o") {
      // Leave Cmd+O for the browser when Open is disabled or unhandled.
      if (!opts.showFileOpen?.() || !opts.onOpenDocument) return;
      e.preventDefault();
      opts.onOpenDocument();
      return;
    }
    if (opts.disableFindReplaceShortcuts?.() && (e.key === "f" || e.key === "h")) return;
    if (e.key === "f" || e.key === "h") {
      e.preventDefault();
      opts.showFindReplace.value = true;
    } else if (e.key === "k") {
      e.preventDefault();
      opts.showHyperlink.value = true;
    }
  }

  onMounted(() => window.addEventListener("keydown", handleKeyDown));
  onBeforeUnmount(() => window.removeEventListener("keydown", handleKeyDown));

  return { handleKeyDown };
}
