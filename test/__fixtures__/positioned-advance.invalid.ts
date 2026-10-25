import { measureRun as rawRun, measureText, measureTextWidth } from "../../packages/core/src/layout-engine/measure/measureProvider";
import * as measure from "../../packages/core/src/layout-engine/measure/measureProvider";

export { rawRun, measure, measureText, measureTextWidth };

function computeLinePmRange() {
  const computeLinePmRange = () => ({ pmStart: 0, pmEnd: 0 });
  return computeLinePmRange();
}

export { computeLinePmRange };
