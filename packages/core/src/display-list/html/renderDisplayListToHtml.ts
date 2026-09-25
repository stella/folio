/**
 * A display list as one HTML document, with no browser and no DOM.
 *
 * The DOM backend paints into whatever `Document` it is handed. This module
 * hands it a document with only what the backend reaches for and serializes
 * the result, so a server, a CLI or a harness can produce the page a browser
 * would show from the same painter the editor uses. Binary data (images,
 * embedded faces) is inlined as `data:` URLs, because a `blob:` URL minted in
 * this process resolves to nothing once the markup leaves it.
 */

import { renderDisplayListToDom } from "../dom/renderDisplayListToDom";
import type { DisplayColor, DisplayList, DisplayPage } from "../types";

type StubElement = {
  readonly tagName: string;
  readonly style: Record<string, string>;
  readonly dataset: Record<string, string>;
  readonly attributes: Record<string, string>;
  readonly children: StubElement[];
  className: string;
  id: string;
  textContent: string;
  alt: string;
  src: string;
  href: string;
  title: string;
  readonly append: (...nodes: StubElement[]) => void;
  readonly appendChild: (node: StubElement) => StubElement;
  readonly setAttribute: (name: string, value: string) => void;
};

const VOID_TAGS = new Set(["img", "br", "hr"]);

const createStubElement = (tagName: string): StubElement => {
  const children: StubElement[] = [];
  const attributes: Record<string, string> = {};
  return {
    tagName,
    style: {},
    dataset: {},
    attributes,
    children,
    className: "",
    id: "",
    textContent: "",
    alt: "",
    src: "",
    href: "",
    title: "",
    append: (...nodes) => {
      children.push(...nodes);
    },
    appendChild: (node) => {
      children.push(node);
      return node;
    },
    setAttribute: (name, value) => {
      attributes[name] = value;
    },
  };
};

/**
 * A `Document` with only what the DOM backend reaches for. A backend that
 * grows a new requirement throws on the missing property here rather than
 * silently serializing less than it painted.
 */
const stubDocument = (): Document => {
  const doc = { createElement: createStubElement };
  // SAFETY: the backend calls `createElement` and then touches `style`,
  // `dataset`, `className`, `id`, `textContent`, `alt`, `src`, `href`, `title`,
  // `append` and `setAttribute`, every one of which `createStubElement` provides.
  return doc as unknown as Document;
};

// SAFETY: only elements created by `stubDocument()` reach this.
const asStub = (element: HTMLElement) => element as unknown as StubElement;

/** A vendor property is camelCase with no leading capital, so its leading dash is explicit. */
const VENDOR_PREFIXES = ["webkit-", "moz-", "ms-", "o-"] as const;

const toKebabCase = (name: string): string => {
  const hyphenated = name.replaceAll(/[A-Z]/gu, (letter) => `-${letter.toLowerCase()}`);
  return VENDOR_PREFIXES.some((prefix) => hyphenated.startsWith(prefix))
    ? `-${hyphenated}`
    : hyphenated;
};

export const escapeHtmlText = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const escapeAttribute = (value: string): string => escapeHtmlText(value).replaceAll('"', "&quot;");

const serializeStyle = (style: Record<string, string>): string =>
  Object.entries(style)
    .map(([property, value]) => `${toKebabCase(property)}: ${value}`)
    .join("; ");

/**
 * Serialized without one character of added whitespace: the backend sets
 * `white-space: pre` on every run, so an indentation newline would be painted.
 */
const RAW_TEXT_TAGS = new Set(["style", "script"]);

const serializeElement = (element: StubElement): string => {
  const attributes: string[] = [];
  if (element.className !== "") attributes.push(`class="${escapeAttribute(element.className)}"`);
  if (element.id !== "") attributes.push(`id="${escapeAttribute(element.id)}"`);
  if (element.href !== "") attributes.push(`href="${escapeAttribute(element.href)}"`);
  if (element.src !== "") attributes.push(`src="${escapeAttribute(element.src)}"`);
  if (element.title !== "") attributes.push(`title="${escapeAttribute(element.title)}"`);
  if (element.tagName === "img") attributes.push(`alt="${escapeAttribute(element.alt)}"`);
  for (const [name, value] of Object.entries(element.attributes)) {
    attributes.push(`${name}="${escapeAttribute(value)}"`);
  }
  for (const [key, value] of Object.entries(element.dataset)) {
    attributes.push(`data-${toKebabCase(key)}="${escapeAttribute(value)}"`);
  }
  const style = serializeStyle(element.style);
  if (style !== "") attributes.push(`style="${escapeAttribute(style)}"`);

  const open = [element.tagName, ...attributes].join(" ");
  if (VOID_TAGS.has(element.tagName)) {
    return `<${open} />`;
  }
  if (element.children.length > 0) {
    return `<${open}>${element.children.map(serializeElement).join("")}</${element.tagName}>`;
  }
  // A `<style>` holds CSS, not markup: escaping it would corrupt the
  // `@font-face` rules the backend emits for embedded faces. A `</` inside
  // it is neutralized so authored text cannot close the element early.
  const inner = RAW_TEXT_TAGS.has(element.tagName)
    ? element.textContent.replaceAll("</", "<\\/")
    : escapeHtmlText(element.textContent);
  return `<${open}>${inner}</${element.tagName}>`;
};

/**
 * Each page sits in a slot of whole pixels. A page height is fractional (A4 is
 * 1122.52 px), so pages stacked in normal flow would start at fractional
 * offsets; the slot rounds the offset while the page keeps its exact size.
 */
const pageSlot = (markup: string, page: DisplayPage): string =>
  `<div class="layout-page-slot" style="position: relative; overflow: hidden; width: ${String(Math.ceil(page.widthPx))}px; height: ${String(Math.ceil(page.heightPx))}px">${markup}</div>`;

export type RenderDisplayListToHtmlOptions = {
  /** `@font-face` rules for the families the list names, placed in the head. */
  readonly fontFaceCss?: string;
  /** The document title. */
  readonly title?: string;
  /** Space between pages. Defaults to 0: pages stack flush. */
  readonly pageGapPx?: number;
  /** CSS color behind the pages. Defaults to white. */
  readonly canvasColor?: string;
  /** Page background, painted before any primitive. */
  readonly pageBackground?: DisplayColor;
};

/** The pages' markup, one slot per display page, in page order. */
export const renderDisplayListPagesToHtml = (
  list: DisplayList,
  options: Pick<RenderDisplayListToHtmlOptions, "pageBackground"> = {},
): string[] =>
  renderDisplayListToDom(list, {
    doc: stubDocument(),
    binaryUrls: "dataUrl",
    ...(options.pageBackground !== undefined && { pageBackground: options.pageBackground }),
  }).map((element, index) => {
    const markup = serializeElement(asStub(element));
    const page = list.pages.at(index);
    return page === undefined ? markup : pageSlot(markup, page);
  });

/** A complete HTML document showing every page of the list. */
export const renderDisplayListToHtml = (
  list: DisplayList,
  options: RenderDisplayListToHtmlOptions = {},
): string => {
  const gap = options.pageGapPx ?? 0;
  const canvas = options.canvasColor ?? "#fff";
  const layout =
    gap > 0
      ? `body { display: flex; flex-direction: column; align-items: center; gap: ${String(gap)}px; padding: ${String(gap)}px 0; }`
      : "";
  return [
    "<!doctype html>",
    `<html><head><meta charset="utf-8"><title>${escapeHtmlText(options.title ?? "")}</title><style>`,
    (options.fontFaceCss ?? "").replaceAll("</", "<\\/"),
    `html, body { margin: 0; padding: 0; background: ${canvas}; }`,
    layout,
    "</style></head><body>",
    renderDisplayListPagesToHtml(list, options).join(""),
    "</body></html>",
  ].join("\n");
};
