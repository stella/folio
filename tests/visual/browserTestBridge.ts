import { installInterleavingBridge } from "./interleavingBridge";
import { resolvePaintedDragTarget, resolvePaintedTableTarget } from "./browserPaintedTargets";

/** Bundled with the playground so browser tests also work against built previews. */
export const browserTestBridge = {
  installInterleavingBridge,
  resolvePaintedDragTarget,
  resolvePaintedTableTarget,
};

declare global {
  var __folioBrowserTestBridge: typeof browserTestBridge | undefined;
}
