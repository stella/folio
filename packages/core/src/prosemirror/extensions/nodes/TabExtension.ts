/**
 * Tab Extension — inline tab character node
 */

import type { PositionalTab } from "../../../types/document";
import { expectTabAttrs } from "../../attrs";
import { createNodeExtension } from "../create";

export const TabExtension = createNodeExtension({
  name: "tab",
  schemaNodeName: "tab",
  nodeSpec: {
    inline: true,
    group: "inline",
    attrs: {
      positional: { default: null },
    },
    selectable: false,
    parseDOM: [
      {
        tag: "span.docx-tab",
        getAttrs(node) {
          if (!(node instanceof HTMLElement) || !node.dataset["docxPositionalTab"]) {
            return null;
          }

          return {
            positional: parsePositionalTab(node),
          };
        },
      },
    ],
    toDOM(node) {
      const { positional } = expectTabAttrs(node);
      const attrs: Record<string, string> = {
        class: "docx-tab",
        style: "display: inline-block; min-width: 16px; white-space: pre;",
      };
      if (positional) {
        attrs["data-docx-positional-tab"] = "true";
        if (positional.relativeTo) {
          attrs["data-docx-tab-relative-to"] = positional.relativeTo;
        }
        if (positional.alignment) {
          attrs["data-docx-tab-alignment"] = positional.alignment;
        }
        if (positional.leader) {
          attrs["data-docx-tab-leader"] = positional.leader;
        }
      }

      return ["span", attrs, "\t"];
    },
  },
});

const parsePositionalTab = (element: HTMLElement): PositionalTab => {
  const relativeTo = element.dataset["docxTabRelativeTo"];
  const alignment = element.dataset["docxTabAlignment"];
  const leader = element.dataset["docxTabLeader"];

  return {
    ...(relativeTo === "margin" || relativeTo === "indent" ? { relativeTo } : {}),
    ...(alignment === "left" || alignment === "center" || alignment === "right"
      ? { alignment }
      : {}),
    ...(leader === "none" ||
    leader === "dot" ||
    leader === "hyphen" ||
    leader === "underscore" ||
    leader === "middleDot"
      ? { leader }
      : {}),
  };
};
