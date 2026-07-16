/** Thin React renderer for the shared content-control widget controller. */

import { useCallback, useEffect, useMemo, useSyncExternalStore } from "react";
import { createPortal } from "react-dom";

import type { EditorView } from "prosemirror-view";
import { useTranslations } from "use-intl";

import { ContentControlWidgetController } from "@stll/folio-core/controller/contentControlWidgetController";
import type { ContentControlListItem } from "@stll/folio-core/controller/contentControlWidgetController";
import { useFolioUI } from "../ui/folio-ui";

type FolioMenuItem = ReturnType<typeof useFolioUI>["Menu"]["Item"];

type ContentControlWidgetsOverlayProps = {
  getEventRoot: () => HTMLElement | null;
  getEditorView: () => EditorView | null;
};

export function ContentControlWidgetsOverlay({
  getEventRoot,
  getEditorView,
}: ContentControlWidgetsOverlayProps) {
  const folioUI = useFolioUI();
  const { Root: Menu, Item: MenuItem } = folioUI.Menu;
  const DatePickerPopover = folioUI.DatePickerPopover;
  const t = useTranslations("folio");
  const controller = useMemo(() => new ContentControlWidgetController(), []);
  const snapshot = useSyncExternalStore(
    controller.subscribe,
    controller.getSnapshot,
    controller.getSnapshot,
  );
  const view = getEditorView();

  useEffect(() => {
    controller.bind(view, getEventRoot());
    return () => controller.destroy();
  }, [controller, getEventRoot, view]);

  const onDropdownPick = useCallback(
    (value: string) => controller.pickDropdown(value),
    [controller],
  );
  const onDateChange = useCallback(
    (value: string | null) => (value ? controller.pickDate(value) : false),
    [controller],
  );

  if (snapshot.status === "closed") {
    return null;
  }

  const style = {
    position: "fixed" as const,
    top: `${snapshot.position.y}px`,
    insetInlineStart: `${snapshot.position.x}px`,
    zIndex: 9999,
  };

  return createPortal(
    <div className="folio-root" style={{ display: "contents" }}>
      <div
        role="dialog"
        aria-label={
          snapshot.status === "dropdown"
            ? t("contentControlDropdownAriaLabel")
            : t("contentControlDateAriaLabel")
        }
        style={style}
        className="bg-popover text-popover-foreground min-w-[10rem] rounded-md border p-1 shadow-md"
      >
        {snapshot.status === "dropdown" && (
          <Menu>
            <div role="menu" className="flex flex-col">
              {snapshot.items.length === 0 ? (
                <div className="text-muted-foreground px-2 py-1 text-sm">
                  {t("contentControlDropdownNoOptions")}
                </div>
              ) : (
                snapshot.items.map((item) => (
                  <ContentControlDropdownItem
                    key={`${item.value}::${item.displayText}`}
                    item={item}
                    MenuItem={MenuItem}
                    onPick={onDropdownPick}
                  />
                ))
              )}
            </div>
          </Menu>
        )}
        {snapshot.status === "date" && (
          <DatePickerPopover
            clearLabel={t("clearDate")}
            defaultOpen
            onChange={onDateChange}
            showIcon={false}
            value={snapshot.currentValue}
          />
        )}
      </div>
    </div>,
    document.body,
  );
}

type ContentControlDropdownItemProps = {
  item: ContentControlListItem;
  MenuItem: FolioMenuItem;
  onPick: (value: string) => boolean;
};

function ContentControlDropdownItem({ item, MenuItem, onPick }: ContentControlDropdownItemProps) {
  const handleClick = useCallback(() => onPick(item.value), [item.value, onPick]);

  return <MenuItem onClick={handleClick}>{item.displayText}</MenuItem>;
}
