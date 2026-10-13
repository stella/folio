import { transform } from "lightningcss";
import type { Selector } from "lightningcss";

const hasEditorRoot = (selector: Selector): boolean =>
  selector.some((component) => {
    if (component.type === "class") return component.name === "folio-root";
    if (component.type !== "pseudo-class") return false;
    if (component.kind !== "where" && component.kind !== "is") return false;
    return component.selectors.length > 0 && component.selectors.every(hasEditorRoot);
  });

const hasGenericEditorSelector = (selector: Selector): boolean =>
  selector.some((component) => {
    if (component.type === "class") return component.name.startsWith("ProseMirror");
    if (component.type === "attribute") return component.name === "contenteditable";
    if (component.type !== "pseudo-class") return false;
    if (component.kind !== "where" && component.kind !== "is" && component.kind !== "not")
      return false;
    return component.selectors.some(hasGenericEditorSelector);
  });

/** Confine generic editor rules while leaving font faces and portaled chrome intact. */
export const scopeEditorCss = (css: string): string => {
  let changed = false;
  const { code } = transform({
    filename: "editor.css",
    code: Buffer.from(css),
    visitor: {
      Selector(selector) {
        if (hasEditorRoot(selector) || !hasGenericEditorSelector(selector)) return selector;
        changed = true;
        return [
          {
            type: "pseudo-class",
            kind: "where",
            selectors: [[{ type: "class", name: "folio-root" }]],
          },
          { type: "combinator", value: "descendant" },
          ...selector,
        ] satisfies Selector;
      },
    },
  });
  return changed ? new TextDecoder().decode(code) : css;
};
