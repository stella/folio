/**
 * Compatibility entry point for wrap types.
 *
 * The portable layout model owns these values. Keep this re-export while
 * consumers migrate from the historical DOCX subpath.
 */
export { isFloatingWrapType, isWrapNone, wrapsAroundText, type WrapType } from "../types/wrap";
