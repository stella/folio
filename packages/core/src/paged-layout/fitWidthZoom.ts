/**
 * fitWidthZoom
 *
 * Framework-agnostic fit-width math for the React `useZoomAndPageInfo` hook.
 *
 * The fit reads the container's border-box width minus the widest
 * non-content edge (borders plus a classic vertical scrollbar) seen so far,
 * never `clientWidth` alone. Whether a classic scrollbar shows depends on the
 * fitted zoom, so a fit that follows `clientWidth` can toggle the scrollbar
 * and with it its own input. Reserving the widest edge makes the basis
 * independent of the scrollbar: a container re-fits once when its gutter first
 * appears and then holds.
 */

// Breathing room kept between the fitted page and the scroll container edges so
// the page does not sit flush against them (matched to the viewport padding).
export const FIT_TO_WIDTH_PADDING_PX = 16;
export const FIT_TO_WIDTH_MIN_ZOOM = 0.1;
// A page never grows past 100% in fit-width: a container wider than the page
// leaves the page at natural size, centred, rather than magnified.
export const FIT_TO_WIDTH_MAX_ZOOM = 1;

type FitWidthZoomParams = {
  /** Border-box width of the scroll container (`offsetWidth`). */
  offsetWidth: number;
  /** Content-box width of the scroll container (`clientWidth`). */
  clientWidth: number;
  /** Natural width of the laid-out page at zoom 1. */
  pageWidth: number;
  /** The widest edge returned by earlier fits of the same container, or 0. */
  reservedEdge: number;
};

type FitWidthZoomResult = {
  /** The fitted zoom, or `null` when the container has no width to fit. */
  zoom: number | null;
  /** Pass back as `reservedEdge` on the next fit of the same container. */
  reservedEdge: number;
};

export const fitWidthZoom = ({
  offsetWidth,
  clientWidth,
  pageWidth,
  reservedEdge,
}: FitWidthZoomParams): FitWidthZoomResult => {
  const edge = Math.max(reservedEdge, offsetWidth - clientWidth);
  const available = offsetWidth - edge - FIT_TO_WIDTH_PADDING_PX;
  if (available <= 0 || pageWidth <= 0) {
    return { zoom: null, reservedEdge: edge };
  }
  const zoom = Math.min(
    FIT_TO_WIDTH_MAX_ZOOM,
    Math.max(FIT_TO_WIDTH_MIN_ZOOM, Math.round((available / pageWidth) * 100) / 100),
  );
  return { zoom, reservedEdge: edge };
};
