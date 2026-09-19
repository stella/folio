/**
 * The one owner of XML character escaping for every part folio writes.
 *
 * Two contexts, two functions. A value going into element content takes
 * {@link escapeXmlText}; a value going between the quotes of an attribute
 * takes {@link escapeXmlAttribute}. Nothing else escapes XML by hand: a
 * second escaper always ends up disagreeing with this one about some
 * character, and the disagreement is a package Word refuses to open.
 *
 * Both functions guarantee the same contract: **the result is always legal
 * XML 1.0 character data.** That is what makes the contract checkable — a
 * caller cannot produce a corrupt part by forgetting something.
 *
 * XML 1.0 §2.2 `Char` admits tab, LF, CR, `#x20`-`#xD7FF`, `#xE000`-`#xFFFD`
 * and the astral planes. Everything else — the rest of C0, `#xFFFE`,
 * `#xFFFF`, an unpaired surrogate — cannot appear in a document at all, and
 * cannot be escaped into one either: `&#0;` is as illegal as a raw NUL
 * (§4.1 requires a character reference to refer to a legal `Char`). The only
 * output that is well-formed is one without them, so they are dropped here.
 * Callers that own an input boundary sanitise before this point, with
 * `sanitizeXmlCharacters`, so a user-visible value loses its illegal
 * characters where something can still report it; the drop below is the last
 * line of defence for a boundary nobody has covered yet.
 *
 * The whitespace rules are where the two contexts part:
 *
 * - §3.3.3 normalises tab, LF and CR in an **attribute value** to a space in
 *   every conformant parser, so a literal one does not survive a save. They
 *   are written as character references, which the same step leaves alone.
 * - §2.11 normalises a literal CR (and CRLF) in **content** to LF on read, so
 *   CR is written as a reference there too. Tab and LF survive as themselves.
 *
 * Both contexts escape all five of `& < > " '`. Only `&` and `<` must be
 * escaped in content, but folio has always written the full set, `>` closes
 * the `]]>` hazard of §2.4 without a special case, and one lexical form
 * across every writer is worth more than the bytes it costs. Every spelling
 * here reads back as the character it stands for, so the rule preserves the
 * value whichever context it lands in.
 */

/** XML 1.0 §2.2 `Char`, negated: the characters no XML document may contain. */
const ILLEGAL_XML_CHARACTER_CLASS =
  "[^\\u0009\\u000A\\u000D\\u0020-\\uD7FF\\uE000-\\uFFFD\\u{10000}-\\u{10FFFF}]";

/**
 * Unicode-mode classes match code points, so a well-formed surrogate pair is a
 * single astral code point and cannot match a BMP class: this matches exactly
 * the unpaired surrogates.
 */
const LONE_SURROGATE_PATTERN = /[\uD800-\uDFFF]/gu;

const ILLEGAL_XML_CHARACTER_PATTERN = new RegExp(ILLEGAL_XML_CHARACTER_CLASS, "gu");
const ILLEGAL_XML_CHARACTER_PROBE = new RegExp(ILLEGAL_XML_CHARACTER_CLASS, "u");

/** What an unpaired surrogate becomes at an input boundary: U+FFFD. */
const REPLACEMENT_CHARACTER = "�";

const TEXT_ESCAPES: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;",
  "\r": "&#13;",
};

const ATTRIBUTE_ESCAPES: Record<string, string> = {
  ...TEXT_ESCAPES,
  "\t": "&#9;",
  "\n": "&#10;",
};

// An illegal character reaches the replacer, misses the escape map, and
// becomes the empty string — so one pass both escapes and strips.
const TEXT_PROBE = new RegExp(`[&<>"'\\r]|${ILLEGAL_XML_CHARACTER_CLASS}`, "u");
const TEXT_PATTERN = new RegExp(`[&<>"'\\r]|${ILLEGAL_XML_CHARACTER_CLASS}`, "gu");
const ATTRIBUTE_PROBE = new RegExp(`[&<>"'\\t\\n\\r]|${ILLEGAL_XML_CHARACTER_CLASS}`, "u");
const ATTRIBUTE_PATTERN = new RegExp(`[&<>"'\\t\\n\\r]|${ILLEGAL_XML_CHARACTER_CLASS}`, "gu");

/** Escape a value for XML element content. Tab and LF survive as themselves. */
export const escapeXmlText = (value: string): string =>
  TEXT_PROBE.test(value)
    ? value.replace(TEXT_PATTERN, (character) => TEXT_ESCAPES[character] ?? "")
    : value;

/**
 * Escape a value for an XML attribute value, of either quoting. Tab, LF and CR
 * become character references so attribute-value normalisation cannot flatten
 * them to spaces.
 */
export const escapeXmlAttribute = (value: string): string =>
  ATTRIBUTE_PROBE.test(value)
    ? value.replace(ATTRIBUTE_PATTERN, (character) => ATTRIBUTE_ESCAPES[character] ?? "")
    : value;

/** Whether `value` holds a character XML 1.0 cannot represent at all. */
export const hasIllegalXmlCharacters = (value: string): boolean =>
  ILLEGAL_XML_CHARACTER_PROBE.test(value);

/**
 * Make a value from outside folio representable in XML: an unpaired surrogate
 * becomes U+FFFD, every other character XML cannot hold is dropped.
 *
 * This is the rule for an input boundary — paste, a document operation, a
 * markdown or legal-source compile, a template fill — where the value is still
 * attached to the request that carried it and a warning still has somewhere to
 * go. `escapeXmlText` and `escapeXmlAttribute` apply the same rule again
 * further down, but by then nobody can say which input lost a character.
 */
export const sanitizeXmlCharacters = (value: string): string =>
  hasIllegalXmlCharacters(value)
    ? value
        .replace(LONE_SURROGATE_PATTERN, REPLACEMENT_CHARACTER)
        .replace(ILLEGAL_XML_CHARACTER_PATTERN, "")
    : value;
