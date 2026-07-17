/**
 * Hyperlink composable — owns the popup data ref and the submit /
 * remove handlers wired into the HyperlinkDialog plus the inline
 * HyperlinkPopup navigate / edit / remove actions.
 *
 * The popup edit / remove actions route through our fork's
 * `editHyperlinkAtCursor` / `removeHyperlinkAtCursor` core commands
 * (which own the contiguous-range lookup + transaction), replacing the
 * upstream's `findHyperlinkRangeAt` + hand-built transaction (that helper
 * does not exist in our core).
 */

import { ref, type Ref } from "vue";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  editHyperlinkAtCursor,
  removeHyperlinkAtCursor,
} from "@stll/folio-core/prosemirror/commands/hyperlink";
import { sanitizeExternalUrl } from "@stll/folio-core/utils/urlSecurity";
import type { HyperlinkPopupData } from "../components/ui/hyperlinkPopupTypes";

export type { HyperlinkPopupData };

type CommandFactory = (...args: readonly unknown[]) => Command;

export type UseHyperlinkManagementOptions = {
  editorView: Ref<EditorView | null>;
  getCommands: () => Record<string, CommandFactory>;
};

export type HyperlinkSubmitData = {
  url?: string;
  bookmark?: string;
  displayText: string;
  tooltip: string;
};

export function useHyperlinkManagement(opts: UseHyperlinkManagementOptions) {
  const hyperlinkPopupData = ref<HyperlinkPopupData | null>(null);

  function runCommand(view: EditorView, command: Command): void {
    command(view.state, (tr) => view.dispatch(tr), view);
  }

  function handleHyperlinkSubmit(data: HyperlinkSubmitData) {
    const view = opts.editorView.value;
    if (!view) return;
    const cmds = opts.getCommands();
    const { empty } = view.state.selection;
    const href = data.bookmark ? `#${data.bookmark}` : data.url;
    if (!href) return;

    if (empty && data.displayText) {
      const factory = cmds["insertHyperlink"];
      if (factory) {
        runCommand(view, factory(data.displayText, href, data.tooltip || undefined));
      }
    } else {
      const factory = cmds["setHyperlink"];
      if (factory) {
        runCommand(view, factory(href, data.tooltip || undefined));
      }
    }
    view.focus();
  }

  function handleHyperlinkRemove() {
    const view = opts.editorView.value;
    if (!view) return;
    const factory = opts.getCommands()["removeHyperlink"];
    if (factory) {
      runCommand(view, factory());
    }
    view.focus();
  }

  function handleHyperlinkPopupNavigate(href: string) {
    const safeHref = sanitizeExternalUrl(href);
    if (!safeHref) return;
    window.open(safeHref, "_blank", "noopener,noreferrer");
    hyperlinkPopupData.value = null;
  }

  function handleHyperlinkPopupEdit(displayText: string, href: string) {
    const view = opts.editorView.value;
    if (!view) return;
    editHyperlinkAtCursor(view, { displayText, href });
    hyperlinkPopupData.value = null;
    view.focus();
  }

  function handleHyperlinkPopupRemove() {
    const view = opts.editorView.value;
    if (!view) return;
    removeHyperlinkAtCursor(view, { popupHref: hyperlinkPopupData.value?.href });
    hyperlinkPopupData.value = null;
    view.focus();
  }

  return {
    hyperlinkPopupData,
    handleHyperlinkSubmit,
    handleHyperlinkRemove,
    handleHyperlinkPopupNavigate,
    handleHyperlinkPopupEdit,
    handleHyperlinkPopupRemove,
  };
}
