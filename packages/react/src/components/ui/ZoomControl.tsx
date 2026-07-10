/**
 * Zoom Control
 *
 * A compact dropdown for choosing the editor's document zoom level. Renders the
 * injectable folio Select primitive (so a consumer's design-system Select can
 * override it), mirroring FontSizePicker. Opt-in: folio chrome renders this only
 * when both a zoom value and an onChange handler are supplied (see
 * FormattingBar).
 */

import { useCallback, useMemo } from "react";

import { useFolioUI } from "../../ui/folio-ui";
import { cn } from "../../lib/utils";

// ============================================================================
// TYPES
// ============================================================================

export type ZoomLevel = {
  value: number;
  label: string;
};

export type ZoomControlProps = {
  /** Current zoom (1 = 100%). */
  value?: number;
  /** Called with the chosen zoom level. */
  onChange?: (zoom: number) => void;
  /** Override the preset levels offered in the dropdown. */
  levels?: ZoomLevel[];
  disabled?: boolean;
  className?: string;
  /** Render the trigger at the smaller toolbar-chrome size. */
  compact?: boolean;
};

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_ZOOM_LEVELS: ZoomLevel[] = [
  { value: 0.5, label: "50%" },
  { value: 0.75, label: "75%" },
  { value: 1, label: "100%" },
  { value: 1.25, label: "125%" },
  { value: 1.5, label: "150%" },
  { value: 2, label: "200%" },
];

// ============================================================================
// COMPONENT
// ============================================================================

export function ZoomControl({
  value = 1,
  onChange,
  levels = DEFAULT_ZOOM_LEVELS,
  disabled = false,
  className,
  compact = false,
}: ZoomControlProps) {
  const {
    Root: Select,
    Trigger: SelectTrigger,
    Value: SelectValue,
    Popup: SelectPopup,
    Item: SelectItem,
  } = useFolioUI().Select;

  const matchingLevel = levels.find((level) => Math.abs(level.value - value) < 0.001);
  const displayLabel = matchingLevel ? matchingLevel.label : `${Math.round(value * 100)}%`;
  // Use the matched preset's exact string so the active item highlights even
  // when `value` is a near-preset float (e.g. 0.9999 from continuous zoom).
  const selectValue = matchingLevel ? matchingLevel.value.toString() : value.toString();
  const handleValueChange = useCallback(
    (newValue: string | null) => {
      if (typeof newValue !== "string") {
        return;
      }
      const zoom = Number.parseFloat(newValue);
      if (!Number.isNaN(zoom)) {
        onChange?.(zoom);
      }
    },
    [onChange],
  );
  const triggerStyle = useMemo(
    () => ({ width: compact ? 76 : 80, height: compact ? 28 : 32 }),
    [compact],
  );

  return (
    <Select value={selectValue} onValueChange={handleValueChange} disabled={disabled}>
      <SelectTrigger
        size="sm"
        className={cn(
          "min-h-0 min-w-0 border-transparent bg-transparent text-[var(--doc-text-muted)] tabular-nums shadow-none hover:bg-[var(--doc-primary-light)] hover:text-[var(--doc-text)] data-[pressed]:bg-[var(--doc-primary-light)]",
          compact ? "text-xs" : "text-sm",
          className,
        )}
        style={triggerStyle}
      >
        <SelectValue placeholder="100%">{displayLabel}</SelectValue>
      </SelectTrigger>
      <SelectPopup>
        {levels.map((level) => (
          <SelectItem key={level.value} value={level.value.toString()}>
            {level.label}
          </SelectItem>
        ))}
      </SelectPopup>
    </Select>
  );
}
