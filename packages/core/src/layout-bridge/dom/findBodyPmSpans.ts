import {
  PAINTED_RANGE_SELECTOR,
  PAINTED_START_SELECTOR,
  paintedStartSelector,
} from "../../utils/paintedPositions";

const BODY_SCOPE = ".layout-page-content";

export function findBodyPmSpans(container: ParentNode): HTMLElement[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>(
      `${BODY_SCOPE} span${PAINTED_RANGE_SELECTOR}`,
    ),
  );
}

export function findBodyEmptyRuns(container: ParentNode): HTMLElement[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>(`${BODY_SCOPE} .layout-empty-run`),
  );
}

export function findBodyPmAnchors(container: ParentNode): HTMLElement[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>(
      `${BODY_SCOPE} ${PAINTED_START_SELECTOR}`,
    ),
  );
}

export function findBodyPmAnchor(
  container: ParentNode,
  pmStart: number,
): HTMLElement | null {
  if (!Number.isFinite(pmStart)) {
    return null;
  }
  return container.querySelector<HTMLElement>(
    `${BODY_SCOPE} ${paintedStartSelector(pmStart)}`,
  );
}
