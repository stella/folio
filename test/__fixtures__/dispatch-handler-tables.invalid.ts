/**
 * Handler tables built where the walk runs: written inline, bound to a local,
 * or carried in by an options object the rule cannot read. Each call allocates
 * the table and its closures again, and the lint rule rejects every form.
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

export const parseThroughLocalOptions = (container: XmlElement) => {
  const modelled: string[] = [];
  const options = {
    element: container,
    container: "w:fonts" as const,
    capturePosition: () => modelled.length,
    handlers: { font: CAPTURE },
  };
  return dispatchChildren(options);
};

export const parseThroughSpreadOptions = (container: XmlElement) => {
  const modelled: string[] = [];
  const shared = { handlers: { font: CAPTURE } };
  return dispatchChildren({
    element: container,
    container: "w:fonts",
    capturePosition: () => modelled.length,
    ...shared,
  });
};
