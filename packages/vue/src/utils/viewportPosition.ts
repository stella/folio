type ViewportPositionOptions = {
  clientX: number;
  clientY: number;
  viewport: Pick<DOMRectReadOnly, "left" | "top">;
};

/** The pages viewport is unscaled; only its painted pages carry the zoom transform. */
export const viewportPosition = ({ clientX, clientY, viewport }: ViewportPositionOptions) => ({
  left: clientX - viewport.left,
  top: clientY - viewport.top,
});
