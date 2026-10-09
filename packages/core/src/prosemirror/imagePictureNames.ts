import type { NonVisualDrawingNames } from "../types/document";

export const PICTURE_NAME_DOM_ATTRIBUTES = {
  name: "data-picture-name",
  alt: "data-picture-alt",
  title: "data-picture-title",
} as const satisfies Record<keyof NonVisualDrawingNames, string>;

/** Keep picture metadata separate from the containing drawing's accessibility strings. */
export const pictureNamesFromDom = (element: HTMLElement): NonVisualDrawingNames =>
  Object.fromEntries(
    Object.entries(PICTURE_NAME_DOM_ATTRIBUTES).flatMap(([key, attribute]) => {
      const value = element.getAttribute(attribute);
      return value === null ? [] : [[key, value]];
    }),
  );

export const pictureNamesDomAttrs = (
  names: NonVisualDrawingNames | undefined,
): Record<string, string> => {
  const values = new Map(Object.entries(names ?? {}));
  return Object.fromEntries(
    Object.entries(PICTURE_NAME_DOM_ATTRIBUTES).flatMap(([key, attribute]) => {
      const value = values.get(key);
      return value === undefined ? [] : [[attribute, value]];
    }),
  );
};
