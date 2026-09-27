/**
 * Toolbar toggles for the editor's side panels, pinned at the end of the
 * toolbar so they never collapse into its overflow menu:
 *
 * - the outline, shown only while the outline is a drawer (in wider tiers it
 *   has a column or a rail of its own);
 * - the comments, with the number of open threads as a badge.
 */

import { ListTreeIcon, MessageSquareTextIcon } from "lucide-react";
import { useFormatter, useTranslations } from "use-intl";

import { cn } from "../lib/utils";
import { ToolbarButton } from "./Toolbar";

/** Counts above this read as "99+". */
const MAX_BADGE_COUNT = 99;

/** The outline drawer toggle: absent unless the outline is a drawer. */
export type OutlineToggleState = "absent" | "closed" | "open";

export type PanelTogglesProps = {
  outline: OutlineToggleState;
  onToggleOutline: () => void;
  /** Whether the comments are on screen (as a column or an open drawer). */
  commentsShown: boolean;
  /** Open comment threads, shown as a badge when there are any. */
  commentCount: number;
  onToggleComments: () => void;
};

export function PanelToggles({
  outline,
  onToggleOutline,
  commentsShown,
  commentCount,
  onToggleComments,
}: PanelTogglesProps) {
  const t = useTranslations("folio");
  const format = useFormatter();
  const commentsLabel = t("comments.visibility");
  const badge =
    commentCount > MAX_BADGE_COUNT
      ? `${format.number(MAX_BADGE_COUNT)}+`
      : format.number(commentCount);

  return (
    // The toolbar hands focus back to the editor after a press; a panel toggle
    // moves it into the drawer it opens instead.
    // oxlint-disable-next-line jsx-a11y/no-static-element-interactions -- stops the toolbar's refocus; the buttons inside are the controls
    <div
      className="flex shrink-0 items-center gap-0.5"
      onMouseUp={(event) => event.stopPropagation()}
    >
      {outline !== "absent" && (
        <ToolbarButton
          active={outline === "open"}
          ariaLabel={t("editor.showDocumentOutline")}
          onClick={onToggleOutline}
          testId="toolbar-outline-toggle"
          title={t("editor.showDocumentOutline")}
        >
          <ListTreeIcon size={16} />
        </ToolbarButton>
      )}
      <button
        aria-pressed={commentsShown}
        className={cn(
          "flex h-8 min-w-8 shrink-0 items-center justify-center gap-1 rounded-md px-1.5",
          "transition-colors duration-100",
          "text-[var(--doc-text-muted)] hover:bg-[var(--doc-primary-light)] hover:text-[var(--doc-text)]",
          commentsShown && "bg-[var(--doc-primary-light)] text-[var(--doc-text)]",
        )}
        data-testid="toolbar-comments-toggle"
        onClick={onToggleComments}
        onMouseDown={(event) => event.preventDefault()}
        title={commentsLabel}
        type="button"
      >
        <MessageSquareTextIcon aria-hidden="true" size={16} />
        <span className="sr-only">{commentsLabel}</span>
        {commentCount > 0 && (
          <span
            className="min-w-4 rounded-full bg-[var(--doc-primary)] px-1 text-center text-[10px] leading-4 font-semibold text-[var(--primary-foreground,var(--doc-page))] tabular-nums"
            data-testid="toolbar-comments-count"
          >
            {badge}
          </span>
        )}
      </button>
    </div>
  );
}
