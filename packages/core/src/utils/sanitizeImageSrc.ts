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

/**
 * Assign a sanitized source to an `<img>` element. When the source is rejected
 * the `src` attribute is left unset: assigning `""` makes some browsers resolve
 * it against the page URL and fire a spurious request for the current page.
 *
 * Structural parameter type so painter unit tests (which fake the DOM) can
 * exercise it without an `HTMLImageElement`.
 */
export function applySanitizedImageSrc(
  imgEl: Pick<HTMLImageElement, "src">,
  src: string | null | undefined,
): void {
  const safeSrc = sanitizeImageSrc(src);
  if (safeSrc !== undefined) {
    imgEl.src = safeSrc;
  }
}
