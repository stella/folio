import type {
  NoBreakHyphenContent,
  RunContent,
  SoftHyphenContent,
  TextContent,
} from "../../types/document";

type HyphenElementType = Extract<
  RunContent,
  { type: "softHyphen" | "noBreakHyphen" }
>["type"];

/** These source units render as native text while keeping one logical atom each. */
export const HYPHEN_TEXT_CARRIERS: {
  readonly softHyphen: "­";
  readonly noBreakHyphen: "‑";
} = {
  softHyphen: "­",
  noBreakHyphen: "‑",
} as const satisfies Record<HyphenElementType, string>;

/**
 * The mark that tells a carrier written for `w:softHyphen` or
 * `w:noBreakHyphen` apart from the same character authored inside `w:t`.
 *
 * Both render identically, so the text alone cannot say which one the source
 * held; the save leg reads this mark to write each back as it was read.
 */
export const HYPHEN_ELEMENT_MARK_NAME = "hyphenElement";

const isHyphenElementType = (key: string): key is HyphenElementType =>
  Object.hasOwn(HYPHEN_TEXT_CARRIERS, key);

/** The inverse of {@link HYPHEN_TEXT_CARRIERS}, derived so the two cannot drift. */
const HYPHEN_ELEMENT_BY_CARRIER: ReadonlyMap<string, HyphenElementType> =
  new Map(
    Object.keys(HYPHEN_TEXT_CARRIERS)
      .filter(isHyphenElementType)
      .map((type): [string, HyphenElementType] => [
        HYPHEN_TEXT_CARRIERS[type],
        type,
      ]),
  );

type HyphenElementRunContent =
  TextContent | SoftHyphenContent | NoBreakHyphenContent;

/**
 * The run content a leaf carrying {@link HYPHEN_ELEMENT_MARK_NAME} writes back.
 *
 * Each carrier becomes its element again and any other character stays text.
 * Typing does not normally reach a marked leaf: the mark is not inclusive and
 * neighbouring carriers hold different ids. Ids only repeat when a copy is pasted
 * beside its source or two paragraphs join; a character typed exactly between
 * two such carriers takes the mark, so a letter stays text here and a typed
 * U+00AD or U+2011 saves as the element.
 */
export const hyphenElementRunContent = (
  text: string,
): HyphenElementRunContent[] => {
  const content: HyphenElementRunContent[] = [];
  let pending = "";
  for (const char of text) {
    const type = HYPHEN_ELEMENT_BY_CARRIER.get(char);
    if (type === undefined) {
      pending += char;
      continue;
    }
    if (pending.length > 0) {
      content.push({ type: "text", text: pending });
      pending = "";
    }
    content.push({ type });
  }
  if (pending.length > 0) {
    content.push({ type: "text", text: pending });
  }
  return content;
};
