import { ptToPx } from "../packages/core/src/layout-engine/measure/measureHelpers";
import { setMeasureProvider } from "../packages/core/src/layout-engine/measure/measureProvider";

/** Deterministic pagination without a browser or installed fonts. */
const FIXED_ADVANCE_RATIO = 0.5;

export const installFixedWidthProvider = (): void => {
  const metricsOf = (fontSize: number | undefined) => {
    const px = ptToPx(fontSize ?? 11);
    return {
      fontSize: fontSize ?? 11,
      ascent: px * 0.8,
      descent: px * 0.2,
      fontBoxAscent: px * 0.9,
      fontBoxDescent: px * 0.25,
      lineHeight: px,
      fontFamily: "fixed",
      singleLineRatio: 1.15,
    };
  };
  const widthOf = (text: string, fontSize: number | undefined) =>
    [...text].length * ptToPx(fontSize ?? 11) * FIXED_ADVANCE_RATIO;

  setMeasureProvider({
    getFontMetrics: (style) => metricsOf(style.fontSize),
    measureTextWidth: (text, style) => widthOf(text, style.fontSize),
    measureText: (text, style) => {
      const metrics = metricsOf(style.fontSize);
      return {
        width: widthOf(text, style.fontSize),
        height: metrics.ascent + metrics.descent,
        ascent: metrics.ascent,
        descent: metrics.descent,
      };
    },
    measureRun: (text, style) => {
      const per = ptToPx(style.fontSize ?? 11) * FIXED_ADVANCE_RATIO;
      const charWidths = [...text].flatMap((char) => (char.length === 2 ? [per, 0] : [per]));
      return {
        width: charWidths.reduce((a, b) => a + b, 0),
        charWidths,
        metrics: metricsOf(style.fontSize),
      };
    },
  });
};
