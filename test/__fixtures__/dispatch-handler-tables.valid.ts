/**
 * The same walk with its tables built once at module scope. What a handler
 * writes into arrives as the walk's `context`, so no closure is allocated per
 * element walked.
 */

import {
  type ChildHandlers,
  type ChildReader,
  dispatchChildrenWithContext,
} from "../../packages/core/src/docx/containerChildren";
import type { XmlElement } from "../../packages/core/src/docx/xmlParser";

const FONT_HANDLERS = {
  font: (_child, modelled) => {
    modelled.push("font");
  },
} as const satisfies ChildHandlers<"w:fonts", string[]>;

const HANDLERS_BY_OWNER = { fonts: FONT_HANDLERS } as const;

const UNDECLARED = {
  AlternateContent: (_child, modelled) => {
    modelled.push("alternate");
  },
} as const satisfies Record<string, ChildReader<string[]>>;

export const parseThing = (container: XmlElement) => {
  const modelled: string[] = [];
  return dispatchChildrenWithContext({
    element: container,
    container: "w:fonts",
    capturePosition: () => modelled.length,
    handlers: HANDLERS_BY_OWNER.fonts,
    undeclared: UNDECLARED,
    context: modelled,
  });
};
