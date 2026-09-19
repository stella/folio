/** Fixture: the same dispatch, read by exhaustion. */

type Content =
  | { type: "run"; text: string }
  | { type: "hyperlink"; href: string }
  | { type: "field"; code: string }
  | { type: "math"; omml: string };

export const describeContent = (content: Content): string => {
  switch (content.type) {
    case "run":
      return content.text;
    case "hyperlink":
      return content.href;
    case "field":
      return content.code;
    case "math":
      return content.omml;
    default: {
      const unsupported: never = content;
      throw new Error(`Unsupported content: ${JSON.stringify(unsupported)}`);
    }
  }
};

/** A chain that ends in a `never` check is exhaustive too. */
export const describeContentByChain = (content: Content): string => {
  if (content.type === "run") {
    return content.text;
  } else if (content.type === "hyperlink") {
    return content.href;
  } else if (content.type === "field") {
    return content.code;
  } else if (content.type === "math") {
    return content.omml;
  } else {
    const unsupported: never = content;
    throw new Error(`Unsupported content: ${JSON.stringify(unsupported)}`);
  }
};

/** Two branches are a pair of cases, not a dispatch table. */
export const isRunOrLink = (content: Content): boolean =>
  content.type === "run" || content.type === "hyperlink";
