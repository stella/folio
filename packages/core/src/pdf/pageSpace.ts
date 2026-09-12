/**
 * The single conversion between display-list space and PDF space.
 *
 * The display list is CSS px at 96 dpi with the origin at the page's top-left
 * and y growing down. PDF user space is points (72/inch) with the origin at
 * the bottom-left and y growing up. Both the scale and the flip happen here
 * and nowhere else: a second place that flips y is a second opinion about
 * where the top of the page is, and the two will eventually disagree.
 *
 * Everything downstream, content streams included, is written in
 * display-list coordinates. The page's base CTM ({@link basePageMatrix}) does
 * the conversion for painted geometry; annotations and destinations, which
 * are not affected by a content stream's CTM, run the same matrix themselves
 * through {@link displayPointToPdf}.
 */

/** 72 points per inch over 96 CSS px per inch. */
export const POINTS_PER_PIXEL = 0.75;

/** `[a b c d e f]` as PDF writes a matrix. */
export type PdfMatrix = readonly [number, number, number, number, number, number];

export const pxToPt = (px: number): number => px * POINTS_PER_PIXEL;

/**
 * Scales px to points and flips y about the page's height, so a primitive
 * emitted at display-list y 0 paints at the top of the page.
 */
export const basePageMatrix = (pageHeightPx: number): PdfMatrix => [
  POINTS_PER_PIXEL,
  0,
  0,
  -POINTS_PER_PIXEL,
  0,
  pxToPt(pageHeightPx),
];

export const applyMatrix = (
  matrix: PdfMatrix,
  x: number,
  y: number,
): { readonly x: number; readonly y: number } => ({
  x: matrix[0] * x + matrix[2] * y + matrix[4],
  y: matrix[1] * x + matrix[3] * y + matrix[5],
});

/** A display-list page point in PDF default user space. */
export const displayPointToPdf = (
  pageHeightPx: number,
  xPx: number,
  yPx: number,
): { readonly x: number; readonly y: number } =>
  applyMatrix(basePageMatrix(pageHeightPx), xPx, yPx);

/**
 * Rotation about a point, expressed in display-list space. The space is
 * y-down, so the usual `[cos sin -sin cos]` reads as a clockwise rotation on
 * screen, which is what a positive CSS/OOXML angle means.
 */
export const rotationMatrix = (
  degrees: number,
  originXPx: number,
  originYPx: number,
): PdfMatrix => {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return [
    cos,
    sin,
    -sin,
    cos,
    originXPx - originXPx * cos + originYPx * sin,
    originYPx - originXPx * sin - originYPx * cos,
  ];
};

/**
 * A DrawingML flip and rotation about one display-list point. CSS applies the
 * rightmost scale first, then the rotation; this matrix does the same.
 */
export const transformMatrix = ({
  degrees,
  originXPx,
  originYPx,
  scaleX,
  scaleY,
}: {
  readonly degrees: number;
  readonly originXPx: number;
  readonly originYPx: number;
  readonly scaleX?: -1;
  readonly scaleY?: -1;
}): PdfMatrix => {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const horizontalScale = scaleX ?? 1;
  const verticalScale = scaleY ?? 1;
  const a = cos * horizontalScale;
  const b = sin * horizontalScale;
  const c = -sin * verticalScale;
  const d = cos * verticalScale;
  return [
    a,
    b,
    c,
    d,
    originXPx - originXPx * a - originYPx * c,
    originYPx - originXPx * b - originYPx * d,
  ];
};
