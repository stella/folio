/** Run in the browser: compare parsed HTML, including head metadata and attributes. */
export const clipboardHtmlProjection = (html: string): string => {
  if (html === "") return "";
  return new DOMParser().parseFromString(html, "text/html").documentElement.outerHTML;
};
