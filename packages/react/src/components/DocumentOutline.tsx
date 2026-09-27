/**
 * Document outline for the docx editor.
 *
 * Owns the outline's chrome (the column, rail or drawer surface, its header
 * and toggles) and the editor-specific data: heading positions over the full
 * document size (the paged editor virtualises pages). The active heading is
 * tracked once by the editor (`useActiveHeading`) and passed in, so every
 * surface marks the same one. The list of headings itself comes from the
 * injected OutlineRail (the consumer's design system, or folio's minimal
 * default), told which presentation it fills.
 *
 * The editor decides the surface from the width it has (see
 * `panelLayout.ts`): a `column` beside the page, a `rail` of heading ticks
 * that can open the full outline as a drawer, or a `drawer` over the page.
 */

import type React from "react";
import { type RefObject, useCallback, useMemo, useRef } from "react";

import { PanelLeftOpenIcon, XIcon } from "lucide-react";
import { useTranslations } from "use-intl";

import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import { useFolioUI } from "../ui/folio-ui";
import type { OutlineItem } from "../ui/folio-ui";
import { headingId } from "./hooks/useActiveHeading";
import { useDrawerFocus } from "./panelDrawer";
import { PANEL_METRICS } from "./panelLayout";

/** Where the outline sits: its own column, a rail of ticks, or a drawer over the page. */
export type DocumentOutlineSurface = "column" | "rail" | "drawer";

const SURFACE_WIDTH = {
  column: PANEL_METRICS.outlineColumnWidth,
  rail: PANEL_METRICS.outlineRailWidth,
  drawer: PANEL_METRICS.drawerWidth,
} as const satisfies Record<DocumentOutlineSurface, number>;

export type DocumentOutlineProps = {
  headings: HeadingInfo[];
  scrollContainerRef: RefObject<HTMLDivElement | null>;
  /** Total ProseMirror document content size — drives proportional tick
   *  placement when pages are virtualised. */
  docSize: number;
  /** The heading the reader is in (its item id). */
  activeId: string | null;
  /** Jump to the heading with item id `id`. */
  onJump: (id: string) => void;
  surface: DocumentOutlineSurface;
  /** Rail: whether the full outline is open as a drawer. */
  expanded?: boolean;
  /** Rail: open the full outline as a drawer. */
  onExpand?: () => void;
  /** Drawer: close it (after a jump, Escape, or the close button). */
  onClose?: () => void;
};

export const DocumentOutline: React.FC<DocumentOutlineProps> = ({
  headings,
  scrollContainerRef,
  docSize,
  activeId,
  onJump,
  surface,
  expanded,
  onExpand,
  onClose,
}) => {
  const t = useTranslations("folio");
  const OutlineRail = useFolioUI().OutlineRail;
  const navRef = useRef<HTMLElement>(null);
  useDrawerFocus(navRef, surface === "drawer" && onClose ? onClose : null);

  const items = useMemo<OutlineItem[]>(
    () =>
      headings.map((heading) => {
        const item: OutlineItem = {
          id: headingId(heading),
          label: heading.text,
          level: heading.level,
        };
        if (typeof heading.pageNumber === "number") {
          item.meta = String(heading.pageNumber);
        }
        return item;
      }),
    [headings],
  );

  const pctById = useMemo(() => {
    const map = new Map<string, number>();
    if (docSize > 0) {
      for (const heading of headings) {
        map.set(headingId(heading), Math.min(99, Math.max(1, (heading.pmPos / docSize) * 100)));
      }
    }
    return map;
  }, [headings, docSize]);

  const handleJump = useCallback(
    (id: string) => {
      onJump(id);
      if (surface === "drawer") {
        onClose?.();
      }
    },
    [onClose, onJump, surface],
  );

  const resolvePct = useCallback((id: string) => pctById.get(id) ?? null, [pctById]);

  if (headings.length < 2) {
    return null;
  }

  const outlineLabel = t("editor.showDocumentOutline");
  const width = SURFACE_WIDTH[surface];
  const rail = (
    <OutlineRail
      activeId={activeId}
      ariaLabel={outlineLabel}
      items={items}
      onJump={handleJump}
      panelWidth={width}
      presentation={surface === "rail" ? "rail" : "panel"}
      resolvePct={resolvePct}
      scrollContainerRef={scrollContainerRef}
      topOffset={0}
    />
  );

  if (surface === "rail") {
    return (
      <nav
        aria-label={outlineLabel}
        className="folio-outline folio-outline--rail"
        data-testid="folio-outline"
        data-folio-outline-surface="rail"
        style={{ width }}
      >
        <button
          aria-expanded={expanded ?? false}
          aria-label={outlineLabel}
          className="folio-outline-icon-button"
          data-testid="folio-outline-expand"
          onClick={onExpand}
          title={outlineLabel}
          type="button"
        >
          <PanelLeftOpenIcon aria-hidden="true" size={16} />
        </button>
        {rail}
      </nav>
    );
  }

  return (
    <nav
      ref={navRef}
      aria-label={outlineLabel}
      className={`folio-outline folio-outline--${surface}`}
      data-testid="folio-outline"
      data-folio-outline-surface={surface}
      style={{ width }}
    >
      <div className="folio-outline-header">
        <span className="folio-outline-title">{t("editor.outlineTitle")}</span>
        {surface === "drawer" && onClose && (
          <button
            aria-label={t("common.closeDialog")}
            className="folio-outline-icon-button"
            onClick={onClose}
            title={t("common.closeDialog")}
            type="button"
          >
            <XIcon aria-hidden="true" size={16} />
          </button>
        )}
      </div>
      {rail}
    </nav>
  );
};
