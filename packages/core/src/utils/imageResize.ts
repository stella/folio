export const IMAGE_RESIZE_HANDLES = [
  "nw",
  "ne",
  "se",
  "sw",
  "n",
  "s",
  "e",
  "w",
] as const;

export type ImageResizeHandle = (typeof IMAGE_RESIZE_HANDLES)[number];

const MIN_IMAGE_PX = 20;
const MAX_IMAGE_PX = 2000;

type ImageResizeOptions = {
  handle: ImageResizeHandle;
  deltaX: number;
  deltaY: number;
  startWidth: number;
  startHeight: number;
  aspect: "locked" | "free";
};

export const calculateResizedImageDimensions = ({
  handle,
  deltaX,
  deltaY,
  startWidth,
  startHeight,
  aspect,
}: ImageResizeOptions): { width: number; height: number } => {
  const drivesWidth = handle.includes("w") || handle.includes("e");
  const drivesHeight = handle.includes("n") || handle.includes("s");
  const signX = handle.includes("w") ? -1 : 1;
  const signY = handle.includes("n") ? -1 : 1;

  if (drivesWidth && drivesHeight && aspect === "locked") {
    const scaleX =
      startWidth > 0 ? (startWidth + deltaX * signX) / startWidth : 1;
    const scaleY =
      startHeight > 0 ? (startHeight + deltaY * signY) / startHeight : 1;
    const desiredScale =
      Math.abs(scaleX - 1) > Math.abs(scaleY - 1) ? scaleX : scaleY;
    const minScale = Math.max(
      startWidth > 0 ? MIN_IMAGE_PX / startWidth : 0,
      startHeight > 0 ? MIN_IMAGE_PX / startHeight : 0,
    );
    const maxScale = Math.min(
      startWidth > 0 ? MAX_IMAGE_PX / startWidth : Infinity,
      startHeight > 0 ? MAX_IMAGE_PX / startHeight : Infinity,
    );
    const scale = Math.min(maxScale, Math.max(minScale, desiredScale));

    return { width: startWidth * scale, height: startHeight * scale };
  }

  const clamp = (dimension: number) =>
    Math.max(MIN_IMAGE_PX, Math.min(MAX_IMAGE_PX, dimension));

  return {
    width: drivesWidth ? clamp(startWidth + deltaX * signX) : startWidth,
    height: drivesHeight ? clamp(startHeight + deltaY * signY) : startHeight,
  };
};
