// Clipboard MIME types and image-file selection. Kept apart from the paste
// sanitizer so paste handling can use them without loading DOMPurify.

/**
 * Custom MIME type for internal clipboard format
 */
export const INTERNAL_CLIPBOARD_TYPE = "application/x-stella-folio";

/**
 * Standard clipboard MIME types
 */
export const CLIPBOARD_TYPES = {
  HTML: "text/html",
  PLAIN: "text/plain",
} as const;

/**
 * Extract image files from clipboard data (if present).
 */
export type ClipboardDataLike = {
  files?: FileList | readonly File[] | undefined;
  getData?: ((format: string) => string) | undefined;
  items?: DataTransferItemList | readonly DataTransferItem[] | undefined;
  types?: ArrayLike<string> | undefined;
};

export function getClipboardImageFiles(
  clipboardData: ClipboardDataLike | null,
): File[] {
  if (!clipboardData) {
    return [];
  }

  const collectFromItems = (): File[] => {
    const items = clipboardData.items;
    if (!items || items.length === 0) {
      return [];
    }
    const files: File[] = [];
    for (const item of Array.from(items)) {
      if (item.kind !== "file") {
        continue;
      }
      if (!item.type.startsWith("image/")) {
        continue;
      }
      const file = item.getAsFile();
      if (file) {
        files.push(file);
      }
    }
    return files;
  };

  const collectFromFiles = (): File[] => {
    const clipboardFiles = clipboardData.files;
    if (!clipboardFiles || clipboardFiles.length === 0) {
      return [];
    }
    return Array.from(clipboardFiles).filter((file) =>
      file.type.startsWith("image/"),
    );
  };

  // Prefer items when available to avoid duplicate representations between items and files.
  const candidates = collectFromItems();
  const files = candidates.length ? candidates : collectFromFiles();
  if (files.length <= 1) {
    return files;
  }

  const preferredTypes = [
    "image/png",
    "image/jpeg",
    "image/webp",
    "image/gif",
    "image/svg+xml",
    "image/bmp",
    "image/tiff",
  ];

  const pickBest = (group: File[]): File => {
    // SAFETY: pickBest is only called with non-empty groups
    let best = group[0]!;
    let bestRank = preferredTypes.indexOf(best.type);
    if (bestRank < 0) {
      bestRank = Number.MAX_SAFE_INTEGER;
    }

    for (const file of group.slice(1)) {
      let rank = preferredTypes.indexOf(file.type);
      if (rank < 0) {
        rank = Number.MAX_SAFE_INTEGER;
      }

      if (rank < bestRank) {
        best = file;
        bestRank = rank;
        continue;
      }

      if (rank === bestRank && file.size > best.size) {
        best = file;
      }
    }

    return best;
  };

  const groups = new Map<string, File[]>();
  for (const file of files) {
    const rawName = file.name.trim();
    const baseName = rawName
      ? rawName.replace(/\.[^/.]+$/, "").toLowerCase()
      : "";
    const key = baseName || `size:${file.size}`;
    const bucket = groups.get(key);
    if (bucket) {
      bucket.push(file);
    } else {
      groups.set(key, [file]);
    }
  }

  const deduped: File[] = [];
  for (const group of groups.values()) {
    deduped.push(pickBest(group));
  }

  return deduped;
}

const ENGINE_CLIPBOARD_TYPES = [
  INTERNAL_CLIPBOARD_TYPE,
  CLIPBOARD_TYPES.HTML,
  CLIPBOARD_TYPES.PLAIN,
] as const;

const hasEngineClipboardContent = (
  clipboardData: ClipboardDataLike,
): boolean => {
  if (!clipboardData.types || !clipboardData.getData) {
    return false;
  }

  const availableTypes = new Set(Array.from(clipboardData.types));
  for (const type of ENGINE_CLIPBOARD_TYPES) {
    if (!availableTypes.has(type)) {
      continue;
    }
    try {
      if (clipboardData.getData(type) !== "") {
        return true;
      }
    } catch {
      // Clipboard reads can be denied by the browser. Keep the image-file
      // fallback available when no readable editor content can be confirmed.
    }
  }
  return false;
};

/**
 * Return image files only when the clipboard has no readable editor content.
 * Rich HTML, internal payloads, and plain text must stay on the normal paste
 * path; some applications include a rendered image beside that content.
 */
export function getStandaloneClipboardImageFiles(
  clipboardData: ClipboardDataLike | null,
): File[] {
  const imageFiles = getClipboardImageFiles(clipboardData);
  if (imageFiles.length === 0 || !clipboardData) {
    return imageFiles;
  }
  return hasEngineClipboardContent(clipboardData) ? [] : imageFiles;
}
