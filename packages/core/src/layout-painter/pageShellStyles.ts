/**
 * The inline styles of a painted page's shell.
 *
 * Every layout pass applies them to every page shell. The page's text
 * inherits colour and font from the shell, so a property written to another
 * value and back within a pass restyles every line painted on the page, on
 * every keystroke. Each property therefore appears once, with its final
 * value: writing a value the shell already holds changes nothing.
 */

import { DEFAULT_FONT_FAMILY } from "../layout-engine/measure/measureHelpers";
import { resolveFontFamily } from "../utils/fontResolver";

type PageShellStyleProperty =
  | "position"
  | "width"
  | "height"
  | "backgroundColor"
  | "color"
  | "overflow"
  | "border"
  | "boxShadow"
  | "outline"
  | "fontFamily"
  | "fontSize";

export type PageShellStyleOptions = NonNullable<
  Parameters<typeof resolveFontFamily>[1]
> & {
  backgroundColor?: string | undefined;
};

/** The shell's styles for a `width` x `height` page, one value per property. */
export const pageShellStyles = (
  width: number,
  height: number,
  options: PageShellStyleOptions,
): Readonly<Record<PageShellStyleProperty, string>> => ({
  position: "relative",
  width: `${width}px`,
  height: `${height}px`,
  backgroundColor: options.backgroundColor ?? "var(--doc-canvas, #ffffff)",
  color: "var(--doc-canvas-text, #000)",
  overflow: "hidden",
  // Page borders are painted as a separate overlay (see
  // renderPageBorderOverlay) so they can honor OOXML offsetFrom, zOrder and
  // display and remain text-relative.
  border: "none",
  boxShadow: "none",
  outline: "none",
  // The default, 11pt Calibri. The fallback chain matches canvas
  // measurement, so text widths do not drift when Carlito stands in for
  // Calibri.
  fontFamily: resolveFontFamily(DEFAULT_FONT_FAMILY, options).cssFallback,
  // Pixels, as canvas measurement uses (11pt = 11 * 96 / 72 px).
  fontSize: `${(11 * 96) / 72}px`,
});

/** Give `element` the shell styles, each property written once. */
export const applyPageShellStyles = (
  element: HTMLElement,
  width: number,
  height: number,
  options: PageShellStyleOptions,
): void => {
  const styles = pageShellStyles(width, height, options);
  for (const property of Object.keys(styles) as PageShellStyleProperty[])
    element.style[property] = styles[property];
};
