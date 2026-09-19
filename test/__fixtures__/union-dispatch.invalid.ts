/** Fixture: a chain of `else if` over a union tag, with no `never` default. */

type Content =
  | { type: "run"; text: string }
  | { type: "hyperlink"; href: string }
  | { type: "field"; code: string }
  | { type: "math"; omml: string };

export const describeContent = (content: Content): string => {
  if (content.type === "run") {
    return content.text;
  } else if (content.type === "hyperlink") {
    return content.href;
  } else if (content.type === "field") {
    return content.code;
  }
  return "";
};
