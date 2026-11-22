import type { LayoutOptions } from "./types";

const FIRST_WHOLE_HEADER_GROUP_COMPATIBILITY_MODE = 15;

/** Modern documents start a leading header group together in a fresh region. */
export const tableHeaderGroupStartForCompatibility = (
  compatibilityMode: number | undefined,
): NonNullable<LayoutOptions["tableHeaderGroupStart"]> =>
  compatibilityMode !== undefined &&
  compatibilityMode >= FIRST_WHOLE_HEADER_GROUP_COMPATIBILITY_MODE
    ? "keep-together"
    : "allow-breaks";
