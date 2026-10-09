/** Convert browser list markup into the editor's numbered-paragraph model. */

import { panic } from "better-result";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

import { mintListInstance } from "../../../docx/listNumberingInstances";
import { createNumberingMap, type NumberingMap } from "../../../docx/numberingParser";
import { expectParagraphAttrs } from "../../attrs";
import { completeNumberingForDoc } from "../../listInstanceReferences";
import { listItemAttrs } from "../../listNumbering";
import { getPackageNumberingDefinitions } from "../../plugins/documentNumbering";

const isList = (element: Element): boolean =>
  element.localName === "ul" || element.localName === "ol";

const hasListHint = (node: PMNode | null): boolean =>
  node?.type.name === "paragraph" && Boolean(expectParagraphAttrs(node)._pastedHtmlList);

/**
 * Flatten only list containers. Other HTML remains in place for the normal
 * clipboard parser, including tables and inline formatting inside list items.
 */
export const flattenPastedHtmlLists = (html: string): string => {
  if (!/<(?:ul|ol)\b/iu.test(html)) {
    return html;
  }
  const host = document.createElement("div");
  host.innerHTML = html;
  let nextGroup = 0;

  const flatten = (list: Element, level: number): DocumentFragment => {
    const fragment = document.createDocumentFragment();
    const group = ++nextGroup;
    const kind = list.localName === "ul" ? "bullet" : "numbered";
    const rawStart = Number(list.getAttribute("start"));
    const start = Number.isSafeInteger(rawStart) && rawStart > 0 ? rawStart : undefined;

    for (const item of Array.from(list.children)) {
      if (item.localName !== "li") {
        continue;
      }
      let paragraph = document.createElement("p");
      let emitted = false;
      const mark = (element: Element): void => {
        element.setAttribute("data-folio-list-group", String(group));
        element.setAttribute("data-folio-list-level", String(Math.min(level, 8)));
        element.setAttribute("data-folio-list-kind", kind);
        if (start !== undefined) {
          element.setAttribute("data-folio-list-start", String(start));
        }
      };
      const flush = (): void => {
        if (!paragraph.hasChildNodes()) {
          return;
        }
        mark(paragraph);
        fragment.append(paragraph);
        paragraph = document.createElement("p");
        emitted = true;
      };

      for (const child of Array.from(item.childNodes)) {
        if (child instanceof Element && isList(child)) {
          flush();
          fragment.append(flatten(child, level + 1));
        } else if (child instanceof Element && child.localName === "p") {
          flush();
          mark(child);
          fragment.append(child);
          emitted = true;
        } else if (child.nodeType !== Node.TEXT_NODE || child.textContent?.trim()) {
          paragraph.append(child);
        }
      }
      flush();
      if (!emitted) {
        mark(paragraph);
        fragment.append(paragraph);
      }
    }
    return fragment;
  };

  for (const list of Array.from(host.querySelectorAll("ul,ol"))) {
    if (!list.parentNode || list.parentElement?.closest("ul,ol")) {
      continue;
    }
    list.replaceWith(flatten(list, 0));
  }
  return host.innerHTML;
};

/** Resolve clipboard-only list hints against fresh document numbering ids. */
export const numberPastedHtmlLists = (slice: Slice, view: EditorView): Slice => {
  let hasList = false;
  slice.content.descendants((node) => {
    if (hasListHint(node)) {
      hasList = true;
      return false;
    }
    return !hasList;
  });
  if (!hasList) {
    return slice;
  }
  const startsWithList = hasListHint(slice.content.firstChild);
  const endsWithList = hasListHint(slice.content.lastChild);
  const source = completeNumberingForDoc(
    getPackageNumberingDefinitions(view.state) ?? undefined,
    view.state.doc,
  );
  let definitions = source;
  let numbering: NumberingMap | null = null;
  const groups = new Map<number, { numId: number; kind: "bullet" | "numbered" }>();
  let changed = false;

  const convert = (node: PMNode): PMNode => {
    let result = node;
    if (node.type.name === "paragraph") {
      const attrs = expectParagraphAttrs(node);
      const hint = attrs._pastedHtmlList;
      if (hint) {
        let group = groups.get(hint.group);
        if (!group || group.kind !== hint.kind) {
          const minted = mintListInstance(definitions, {
            kind: hint.kind,
            ...(hint.start === undefined ? {} : { start: hint.start }),
          });
          definitions = minted.definitions;
          numbering = createNumberingMap(definitions);
          group = { numId: minted.numId, kind: hint.kind };
          groups.set(hint.group, group);
        }
        if (!numbering) {
          panic("Pasted list group has no numbering definition");
        }
        const { _pastedHtmlList: _consumed, ...unhinted } = attrs;
        result = node.type.create(
          listItemAttrs(unhinted, { numId: group.numId, ilvl: hint.level }, numbering),
          node.content,
          node.marks,
        );
        changed = true;
      }
    }
    if (result.childCount === 0) {
      return result;
    }
    const children: PMNode[] = [];
    result.forEach((child) => children.push(convert(child)));
    return result.copy(Fragment.fromArray(children));
  };

  const children: PMNode[] = [];
  slice.content.forEach((node) => children.push(convert(node)));
  return changed
    ? new Slice(
        Fragment.fromArray(children),
        startsWithList ? 0 : slice.openStart,
        endsWithList ? 0 : slice.openEnd,
      )
    : slice;
};
