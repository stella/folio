/**
 * Two handler tables built where the walk runs: one written inline, one bound
 * to a local first. Either way every call allocates the table and its closures
 * again, and the lint rule rejects both.
 */

import { CAPTURE, dispatchChildren } from "../../packages/core/src/docx/containerChildren";
import type { XmlElement } from "../../packages/core/src/docx/xmlParser";

export const parseInline = (container: XmlElement) => {
  const modelled: string[] = [];
  return dispatchChildren({
    element: container,
    container: "w:fonts",
    capturePosition: () => modelled.length,
    handlers: {
      font: () => {
        modelled.push("font");
      },
    },
  });
};

export const parseThroughLocal = (container: XmlElement) => {
  const modelled: string[] = [];
  const handlers = { font: CAPTURE };
  return dispatchChildren({
    element: container,
    container: "w:fonts",
    capturePosition: () => modelled.length,
    handlers,
  });
};
