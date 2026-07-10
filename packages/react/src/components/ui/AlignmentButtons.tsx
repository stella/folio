/**
 * Alignment dropdown — shows current alignment icon, opens a popover
 * with left/center/right/justify options.
 */

import { useCallback, useMemo } from "react";
import type { CSSProperties, MouseEvent } from "react";

import { panic } from "better-result";
import {
  AlignCenterIcon,
  AlignJustifyIcon,
  AlignLeftIcon,
  AlignRightIcon,
  ChevronDownIcon,
} from "lucide-react";

import type { ParagraphAlignment } from "@stll/folio-core/types/document";
import { cn } from "../../lib/utils";
import { useFolioUI } from "../../ui/folio-ui";

export type AlignmentButtonsProps = {
  value?: ParagraphAlignment;
  onChange?: (alignment: ParagraphAlignment) => void;
  disabled?: boolean;
  className?: string;
  style?: CSSProperties;
};

const ICON_SIZE = 16;

const stopMouseDownPropagation = (event: MouseEvent) => event.stopPropagation();

const OPTIONS = [
  {
    value: "left" as const,
    label: "Align Left",
    shortcut: "Ctrl+L",
    Icon: AlignLeftIcon,
  },
  {
    value: "center" as const,
    label: "Center",
    shortcut: "Ctrl+E",
    Icon: AlignCenterIcon,
  },
  {
    value: "right" as const,
    label: "Align Right",
    shortcut: "Ctrl+R",
    Icon: AlignRightIcon,
  },
  {
    value: "both" as const,
    label: "Justify",
    shortcut: "Ctrl+J",
    Icon: AlignJustifyIcon,
  },
];

export function AlignmentButtons({
  value = "left",
  onChange,
  disabled = false,
}: AlignmentButtonsProps) {
  const {
    Root: Popover,
    Trigger: PopoverTrigger,
    Popup: PopoverPopup,
    Close: PopoverClose,
  } = useFolioUI().Popover;
  const defaultOption = OPTIONS[0];
  if (!defaultOption) {
    panic("AlignmentButtons: OPTIONS is empty");
  }
  const current = OPTIONS.find((o) => o.value === value) ?? defaultOption;

  return (
    <Popover>
      <PopoverTrigger
        className={cn(
          "flex h-8 w-8 items-center justify-center rounded-md transition-colors",
          "text-[var(--doc-text-muted)] hover:bg-[var(--doc-primary-light)] hover:text-[var(--doc-text)]",
          disabled &&
            "cursor-not-allowed text-[var(--doc-text-subtle)] opacity-[0.16] disabled:hover:bg-transparent disabled:hover:text-[var(--doc-text-subtle)]",
        )}
        data-testid="toolbar-alignment"
        disabled={disabled}
      >
        <current.Icon size={ICON_SIZE} />
        <ChevronDownIcon size={12} className="-ms-0.5" />
      </PopoverTrigger>
      <PopoverPopup
        side="bottom"
        sideOffset={4}
        className="flex gap-0.5 p-1"
        onMouseDown={stopMouseDownPropagation}
      >
        {OPTIONS.map((opt) => (
          <AlignmentOption
            key={opt.value}
            PopoverClose={PopoverClose}
            option={opt}
            selected={value === opt.value}
            onChange={onChange}
          />
        ))}
      </PopoverPopup>
    </Popover>
  );
}

type AlignmentOptionValue = (typeof OPTIONS)[number];
type PopoverCloseComponent = ReturnType<typeof useFolioUI>["Popover"]["Close"];

type AlignmentOptionProps = {
  PopoverClose: PopoverCloseComponent;
  option: AlignmentOptionValue;
  selected: boolean;
  onChange: AlignmentButtonsProps["onChange"];
};

function AlignmentOption({ PopoverClose, option, selected, onChange }: AlignmentOptionProps) {
  const handleMouseDown = useCallback((event: MouseEvent) => event.preventDefault(), []);
  const handleClick = useCallback(() => onChange?.(option.value), [onChange, option.value]);
  const trigger = useMemo(
    () => (
      <button
        className={cn(
          "flex size-8 items-center justify-center rounded transition-colors",
          selected
            ? "bg-[var(--doc-primary-light)] text-[var(--doc-text)]"
            : "text-[var(--doc-text)] hover:bg-[var(--doc-primary-light)]",
        )}
        data-testid={`alignment-${option.value}`}
        onMouseDown={handleMouseDown}
        onClick={handleClick}
        title={`${option.label} (${option.shortcut})`}
        type="button"
      />
    ),
    [handleClick, handleMouseDown, option.label, option.shortcut, option.value, selected],
  );

  return (
    <PopoverClose render={trigger}>
      <option.Icon size={ICON_SIZE} />
    </PopoverClose>
  );
}
