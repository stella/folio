const RASTER_SIGNATURES = [
  { mimeType: "image/png", chunks: [{ offset: 0, bytes: [0x89, 0x50, 0x4e, 0x47] }] },
  { mimeType: "image/jpeg", chunks: [{ offset: 0, bytes: [0xff, 0xd8] }] },
  { mimeType: "image/gif", chunks: [{ offset: 0, bytes: [0x47, 0x49, 0x46, 0x38] }] },
  { mimeType: "image/bmp", chunks: [{ offset: 0, bytes: [0x42, 0x4d] }] },
  {
    mimeType: "image/webp",
    chunks: [
      { offset: 0, bytes: [0x52, 0x49, 0x46, 0x46] },
      { offset: 8, bytes: [0x57, 0x45, 0x42, 0x50] },
    ],
  },
  { mimeType: "image/tiff", chunks: [{ offset: 0, bytes: [0x49, 0x49, 0x2a, 0x00] }] },
  { mimeType: "image/tiff", chunks: [{ offset: 0, bytes: [0x4d, 0x4d, 0x00, 0x2a] }] },
  { mimeType: "image/tiff", chunks: [{ offset: 0, bytes: [0x49, 0x49, 0x2b, 0x00] }] },
  { mimeType: "image/tiff", chunks: [{ offset: 0, bytes: [0x4d, 0x4d, 0x00, 0x2b] }] },
] as const;

export type RasterMimeType = (typeof RASTER_SIGNATURES)[number]["mimeType"];

export const RASTER_MIME_TYPES: ReadonlySet<string> = new Set(
  RASTER_SIGNATURES.map(({ mimeType }) => mimeType),
);

export const detectRasterMimeType = (data: ArrayBuffer): RasterMimeType | undefined => {
  const bytes = new Uint8Array(data);
  return RASTER_SIGNATURES.find(({ chunks }) =>
    chunks.every(({ offset, bytes: signature }) =>
      signature.every((byte, index) => bytes[offset + index] === byte),
    ),
  )?.mimeType;
};
