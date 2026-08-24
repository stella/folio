const TABLE_OF_CONTENTS_STYLE_ID = /^TOC(?<level>\d*)$/iu;
const TABLE_OF_CONTENTS_STYLE_NAME = /^toc\s+(?<level>\d+)$/iu;

type TableOfContentsStyleIdentity = {
  styleId: string | undefined;
  styleName?: string;
};

const parsedLevel = (match: RegExpExecArray | null): number | undefined => {
  if (!match) {
    return undefined;
  }
  const rawLevel = match.groups?.["level"];
  if (rawLevel === "") {
    return 1;
  }
  const level = Number(rawLevel);
  return Number.isSafeInteger(level) && level > 0 ? level : undefined;
};

/** Resolve a built-in TOC entry style without relying on its localized style id. */
export const tableOfContentsStyleLevel = ({
  styleId,
  styleName,
}: TableOfContentsStyleIdentity): number | undefined =>
  parsedLevel(TABLE_OF_CONTENTS_STYLE_ID.exec(styleId ?? "")) ??
  parsedLevel(TABLE_OF_CONTENTS_STYLE_NAME.exec(styleName ?? ""));
