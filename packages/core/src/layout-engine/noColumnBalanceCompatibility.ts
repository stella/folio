import type { LayoutOptions } from "./types";

type ColumnBalancing = NonNullable<LayoutOptions["columnBalancing"]>;

/** `w:noColumnBalance` is ignored beginning with compatibility mode 15. */
export const columnBalancingForCompatibility = (
  noColumnBalance: boolean | undefined,
  compatibilityMode: number | undefined,
): ColumnBalancing =>
  noColumnBalance === true &&
  (compatibilityMode === undefined || compatibilityMode < 15)
    ? "natural"
    : "balanced";
