import { insertHyperlink, setHyperlink } from "@stll/folio-core/prosemirror";

import {
  resolvePaintedDragTarget,
  resolvePaintedTableTarget,
} from "./browserPaintedTargets";
import { installInterleavingBridge } from "./interleavingBridge";

/** Bundled with the playground so browser tests also work against built previews. */
export const browserTestBridge = {
  installInterleavingBridge,
  setHyperlink,
  insertHyperlink,
  resolvePaintedDragTarget,
  resolvePaintedTableTarget,
};

declare global {
  var __folioBrowserTestBridge: typeof browserTestBridge | undefined;
}
