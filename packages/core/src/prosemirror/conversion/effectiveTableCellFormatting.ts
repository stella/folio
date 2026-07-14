import type {
  CellMargins,
  ShadingProperties,
  TableBorders,
  TableCellBorders,
  TableCellFormatting,
  TableWidthType,
  Theme,
} from "../../types/document";
import { resolveColor } from "../../utils/colorResolver";
import { resolveShadingFill } from "../../utils/formatToStyle";

export type TableCellMarginsAttrs = {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
};

type EffectiveTableCellWidth =
  | { type: "none" }
  | {
      type: "value";
      source: "direct" | "grid";
      value: number;
      widthType: TableWidthType;
    };

type EffectiveTableCellBackground =
  | { type: "none"; source: "none" | "direct" | "style" }
  | { type: "color"; source: "direct" | "style"; rgb: string };

type EffectiveTableCellFormatting = Readonly<{
  width: EffectiveTableCellWidth;
  background: EffectiveTableCellBackground;
  borders: TableCellBorders | undefined;
  margins: TableCellMarginsAttrs | undefined;
}>;

export type TableCellPosition = {
  isFirstRow?: boolean;
  isLastRow?: boolean;
  isFirstColumn?: boolean;
  isLastColumn?: boolean;
};

type ResolveEffectiveTableCellFormattingOptions = {
  directFormatting: TableCellFormatting | undefined;
  styleFormatting: TableCellFormatting | undefined;
  tableBorders: TableBorders | undefined;
  position: TableCellPosition;
  gridWidthPercent: number | undefined;
  defaultMargins: TableCellMarginsAttrs | undefined;
  theme: Theme | null | undefined;
};

const TABLE_BORDER_SIDES = [
  "top",
  "bottom",
  "left",
  "right",
  "insideH",
  "insideV",
  "topLeftToBottomRight",
  "topRightToBottomLeft",
] as const satisfies readonly (keyof TableCellBorders)[];

/**
 * Resolve the display-facing table-cell cascade without mutating authored
 * OOXML formatting. Source discriminators keep inherited values distinct from
 * direct values that the editor may later override.
 */
export const resolveEffectiveTableCellFormatting = ({
  directFormatting,
  styleFormatting,
  tableBorders,
  position,
  gridWidthPercent,
  defaultMargins,
  theme,
}: ResolveEffectiveTableCellFormattingOptions): EffectiveTableCellFormatting => {
  const width = resolveEffectiveWidth(directFormatting, gridWidthPercent);
  const background = resolveEffectiveBackground({ directFormatting, styleFormatting, theme });
  const borders = resolveEffectiveBorders({
    directFormatting,
    styleFormatting,
    tableBorders,
    position,
    theme,
  });
  const margins = resolveEffectiveMargins({
    directFormatting,
    styleFormatting,
    defaultMargins,
  });

  return { width, background, borders, margins };
};

const resolveEffectiveWidth = (
  directFormatting: TableCellFormatting | undefined,
  gridWidthPercent: number | undefined,
): EffectiveTableCellWidth => {
  const directWidth = directFormatting?.width;
  if (directWidth) {
    return {
      type: "value",
      source: "direct",
      value: directWidth.value,
      widthType: directWidth.type,
    };
  }
  if (gridWidthPercent !== undefined) {
    return { type: "value", source: "grid", value: gridWidthPercent, widthType: "pct" };
  }
  return { type: "none" };
};

type ResolveEffectiveBackgroundOptions = Pick<
  ResolveEffectiveTableCellFormattingOptions,
  "directFormatting" | "styleFormatting" | "theme"
>;

const resolveEffectiveBackground = ({
  directFormatting,
  styleFormatting,
  theme,
}: ResolveEffectiveBackgroundOptions): EffectiveTableCellBackground => {
  const directShading = directFormatting?.shading;
  if (directShading !== undefined) {
    return resolveBackgroundColor({ shading: directShading, source: "direct", theme });
  }

  const styleShading = styleFormatting?.shading;
  if (styleShading !== undefined) {
    return resolveBackgroundColor({ shading: styleShading, source: "style", theme });
  }

  return { type: "none", source: "none" };
};

type ResolveBackgroundColorOptions = {
  shading: ShadingProperties;
  source: "direct" | "style";
  theme: Theme | null | undefined;
};

const resolveBackgroundColor = ({
  shading,
  source,
  theme,
}: ResolveBackgroundColorOptions): EffectiveTableCellBackground => {
  const rgb = resolveShadingFill(shading, theme).replace(/^#/u, "");

  if (!rgb) {
    return { type: "none", source };
  }
  return { type: "color", source, rgb };
};

type ResolveEffectiveBordersOptions = Pick<
  ResolveEffectiveTableCellFormattingOptions,
  "directFormatting" | "styleFormatting" | "tableBorders" | "position" | "theme"
>;

const resolveEffectiveBorders = ({
  directFormatting,
  styleFormatting,
  tableBorders,
  position,
  theme,
}: ResolveEffectiveBordersOptions): TableCellBorders | undefined => {
  const baseBorders = tableBorders
    ? {
        top: position.isFirstRow ? tableBorders.top : tableBorders.insideH,
        bottom: position.isLastRow ? tableBorders.bottom : tableBorders.insideH,
        left: position.isFirstColumn ? tableBorders.left : tableBorders.insideV,
        right: position.isLastColumn ? tableBorders.right : tableBorders.insideV,
      }
    : undefined;
  const styleBorders = styleFormatting?.borders;
  const directBorders = directFormatting?.borders;
  const borders =
    baseBorders || styleBorders || directBorders
      ? { ...baseBorders, ...styleBorders, ...directBorders }
      : undefined;

  return resolveThemedBorderColors(borders, theme);
};

type ResolveEffectiveMarginsOptions = Pick<
  ResolveEffectiveTableCellFormattingOptions,
  "directFormatting" | "styleFormatting" | "defaultMargins"
>;

const resolveEffectiveMargins = ({
  directFormatting,
  styleFormatting,
  defaultMargins,
}: ResolveEffectiveMarginsOptions): TableCellMarginsAttrs | undefined => {
  if (directFormatting?.margins) {
    return cellMarginsToAttrs(directFormatting.margins);
  }
  if (styleFormatting?.margins) {
    return cellMarginsToAttrs(styleFormatting.margins);
  }
  return defaultMargins;
};

const cellMarginsToAttrs = (margins: CellMargins): TableCellMarginsAttrs => {
  const result: TableCellMarginsAttrs = {};
  if (margins.top?.value !== undefined) {
    result.top = margins.top.value;
  }
  if (margins.bottom?.value !== undefined) {
    result.bottom = margins.bottom.value;
  }
  if (margins.left?.value !== undefined) {
    result.left = margins.left.value;
  }
  if (margins.right?.value !== undefined) {
    result.right = margins.right.value;
  }
  return result;
};

const resolveThemedBorderColors = (
  borders: TableCellBorders | undefined,
  theme: Theme | null | undefined,
): TableCellBorders | undefined => {
  if (!borders || !theme?.colorScheme) {
    return borders;
  }

  let resolved: TableCellBorders | undefined;
  for (const side of TABLE_BORDER_SIDES) {
    const border = borders[side];
    if (!border?.color?.themeColor || border.color.auto) {
      continue;
    }

    resolved ??= { ...borders };
    resolved[side] = {
      ...border,
      color: {
        rgb: resolveColor(border.color, theme).replace(/^#/u, ""),
      },
    };
  }

  return resolved ?? borders;
};
