import { Popover as PopoverPrimitive } from "@base-ui/react/popover";
import { useCallback, useMemo } from "react";

import { cn } from "../../lib/utils";
import type { ColorPreset, FolioColorPickerProps } from "../folio-ui";

/**
 * Built-in, dependency-light ColorPicker used when a consumer does not inject
 * one. Wraps `@base-ui/react`'s Popover primitive (focus management,
 * portalling, collision-aware positioning) and renders the `presets` as a row
 * of swatch buttons plus a native `<input type="color">` for a custom color.
 * `onSelect` emits a preset value or a 6-char uppercase hex (no `#`), matching
 * the design-system ColorPicker contract. Consumers inject a polished picker
 * via `DocxEditor`'s `components` prop.
 */
const swatchColor = (preset: ColorPreset): string => preset.color ?? `#${preset.value}`;

const normalizeHex = (hex: string): string => hex.replace(/^#/u, "").toUpperCase();

const EMPTY_PRESETS: ColorPreset[] = [];

export function DefaultColorPicker({
  value,
  onSelect,
  onClear,
  presets = EMPTY_PRESETS,
  columns = 8,
  children,
}: FolioColorPickerProps) {
  const trigger = useMemo(() => <div className="folio-default-color-picker-trigger" />, []);

  return (
    <PopoverPrimitive.Root>
      <PopoverPrimitive.Trigger nativeButton={false} render={trigger}>
        {children}
      </PopoverPrimitive.Trigger>
      <PopoverPrimitive.Portal>
        <PopoverPrimitive.Positioner
          align="start"
          // `folio-root` re-establishes the editor root inside the body portal
          // so design tokens and the standalone stylesheet's scoped utilities
          // apply to the portalled content (and `.dark .folio-root` themes it).
          className="folio-default-popover-positioner folio-root"
          side="bottom"
          sideOffset={4}
        >
          <PopoverPrimitive.Popup className="folio-default-color-picker-popup">
            {onClear && (
              <PopoverPrimitive.Close
                className="folio-default-color-picker-clear"
                onClick={onClear}
              >
                No color
              </PopoverPrimitive.Close>
            )}
            <div
              className="folio-default-color-picker-swatches"
              style={{ gridTemplateColumns: `repeat(${columns}, 1fr)` }}
            >
              {presets.map((preset) => (
                <ColorSwatch
                  key={preset.value}
                  onSelect={onSelect}
                  preset={preset}
                  selected={value === preset.value}
                />
              ))}
            </div>
            <label className="folio-default-color-picker-custom">
              Custom
              <input
                aria-label="Custom color"
                onChange={(event) => onSelect?.(normalizeHex(event.target.value))}
                type="color"
                value={value ? `#${normalizeHex(value)}` : "#000000"}
              />
            </label>
          </PopoverPrimitive.Popup>
        </PopoverPrimitive.Positioner>
      </PopoverPrimitive.Portal>
    </PopoverPrimitive.Root>
  );
}

type ColorSwatchProps = {
  onSelect: FolioColorPickerProps["onSelect"];
  preset: ColorPreset;
  selected: boolean;
};

function ColorSwatch({ onSelect, preset, selected }: ColorSwatchProps) {
  const handleClick = useCallback(() => onSelect?.(preset.value), [onSelect, preset.value]);
  const style = useMemo(() => ({ backgroundColor: swatchColor(preset) }), [preset]);

  return (
    <PopoverPrimitive.Close
      aria-label={preset.label}
      className={cn(
        "folio-default-color-picker-swatch",
        selected && "folio-default-color-picker-swatch--selected",
      )}
      onClick={handleClick}
      style={style}
      title={preset.label}
    />
  );
}
