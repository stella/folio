/**
 * Keep painted image sources local to bytes the host already owns.
 *
 * Document media and pasted files enter the editor as `data:image/*` or
 * `blob:` URLs. Network and executable schemes are rejected so opening or
 * pasting an untrusted document cannot trigger an external request or run a
 * script scheme through `<img src>`.
 */
export function sanitizeImageSrc(src: string | null | undefined): string | undefined {
  if (typeof src !== "string") {
    return undefined;
  }
  const value = src.trim();
  if (!value) {
    return undefined;
  }

  if (/^blob:/iu.test(value)) {
    return value;
  }
  if (/^data:image\/[a-z0-9.+-]+(?:;[^,]*)?,/iu.test(value)) {
    return value;
  }

  return undefined;
}
