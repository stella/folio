/**
 * Shared parser for Structured Document Tag properties (`w:sdtPr`).
 *
 * One parser feeds the inline (run-level), block-level, row and cell SDT paths
 * so none of them can drift. `CT_SdtPr` goes through the shared child
 * dispatcher: the handler map is total over the children the schema declares
 * (ECMA-376 §17.5.2.38), so a child folio does not model reaches
 * {@link SdtProperties.preserved} instead of falling off the end of a
 * `switch`, and anything undeclared — `w14:checkbox`, `w15:appearance`,
 * `w15:color`, `w15:repeatingSection`, an element a later revision adds — goes
 * there too, because the extension namespaces are not in the Transitional
 * graph at all.
 *
 * `CT_SdtPr` is an `xsd:sequence`, so the sink records the schema ordinal
 * rather than a count of modelled siblings and
 * {@link serializeSdtProperties} merges the two halves back in that order.
 */

import type { PreservedMarkup, SdtProperties } from "../types/document";
import {
  CAPTURE,
  type ChildHandlers,
  dispatchChildren,
  keptUnless,
  sequencePositions,
} from "./containerChildren";
import type { DeclaredChild } from "./containerChildren.gen";
import { SdtLockSchema, narrowEnum } from "./parserEnums";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChild,
  getAttributeAnyPrefix,
  parseBooleanElement,
  type XmlElement,
} from "./xmlParser";

function parseListItems(el: XmlElement): { displayText: string; value: string }[] {
  const items: { displayText: string; value: string }[] = [];
  for (const child of el.elements ?? []) {
    if (
      child.type === "element" &&
      (child.name === "w:listItem" || child.name?.endsWith(":listItem"))
    ) {
      // OOXML §17.5.2.10: `w:displayText` is optional, and `w:value` is
      // optional too. Fall back each to the other so a partially specified
      // listItem stays selectable + visible — without this, the dropdown
      // shell renders a blank option and the value-set path writes an
      // empty paragraph for a real OOXML pick.
      //
      // Read attributes by the element's own prefix first (the spec lets
      // a producer bind the Word namespace under any prefix, e.g.
      // `<ns0:listItem ns0:displayText="A" ns0:value="a"/>`). The previous
      // hard-coded `w:` lookup turned every such item into `(null, null)`
      // and silently dropped it, so non-standard-prefix dropdowns opened
      // with no options.
      const displayText = getAttributeAnyPrefix(child, "displayText");
      const value = getAttributeAnyPrefix(child, "value");
      if (displayText === null && value === null) {
        continue;
      }
      items.push({
        displayText: displayText ?? value ?? "",
        value: value ?? displayText ?? "",
      });
    }
  }
  return items;
}

/**
 * Local names that always belong to one specific OOXML namespace. The
 * serializer's document root only declares the canonical `w` / `w14` /
 * `w15` prefixes, so any captured raw SDT child written under an
 * alternate prefix would replay as an undefined-prefix element (Word
 * refuses such files). We rewrite by element-local-name so an inherited
 * `<x:checkbox>` / `<y:repeatingSection>` (where `x` / `y` are bound to
 * the w14 / w15 URIs at the source's document root) always lands under
 * the canonical prefix in our output regardless of the source's prefix
 * choice.
 */
// Only local names that are UNAMBIGUOUSLY bound to one namespace make it
// in here. `color` and `appearance` would otherwise look identical to
// `<w:color>` inside a placeholder `<w:rPr>` and we'd silently rewrite
// run formatting as w15 SDT appearance, corrupting the parse → save
// round trip for any sdtPr that carries nested run properties.
//
// W_LOCAL_NAMES covers SDT property children that exist only in the
// WordprocessingML (w:) namespace by spec — `<x:tag>` inherited from
// the document root must mean `<w:tag>` because no other namespace
// defines that local name in a `<w:sdtPr>`/`<w:sdtEndPr>` context.
const W_LOCAL_NAMES = new Set([
  "alias",
  "tag",
  "id",
  "lock",
  "placeholder",
  "docPart",
  "showingPlcHdr",
  "text",
  "date",
  "dateFormat",
  "lid",
  "calendar",
  "storeMappedDataAs",
  "dropDownList",
  "listItem",
  "comboBox",
  "picture",
  "docPartObj",
  "docPartList",
  "docPartCategory",
  "docPartGallery",
  "docPartUnique",
  "group",
  "equation",
  "citation",
  "bibliography",
  "richText",
]);
const W14_LOCAL_NAMES = new Set(["checkbox", "checked"]);
const W15_LOCAL_NAMES = new Set(["repeatingSection", "repeatingSectionItem"]);

/**
 * Rewrite the captured raw `<*:sdtPr>` / `<*:sdtEndPr>` snippet so every
 * SDT-namespace element uses the canonical `w:` / `w14:` / `w15:`
 * prefix on save. The blockSdtSerializer's document root declares only
 * those three prefixes, so replaying a source snippet that uses an
 * alternate prefix (`<ns0:sdtPr>` with `xmlns:ns0` declared on the
 * source's `<w:document>`, or `<x:checkbox>` inside a canonical sdtPr)
 * would produce invalid XML in the saved DOCX — Word refuses files
 * with unresolved namespace prefixes.
 *
 * Heuristics, since fast-xml-parser does not surface namespace URIs in
 * preserveOrder mode:
 *
 * 1. If the captured wrapper element itself uses a non-`w` prefix, that
 *    prefix is taken to be bound to the WP URI (the source DOCX would
 *    otherwise be invalid) and ALL occurrences of it inside the snippet
 *    are rewritten to `w:`.
 * 2. After step 1, any remaining alt-prefix on an element whose local
 *    name is in W14_LOCAL_NAMES / W15_LOCAL_NAMES gets normalized to
 *    `w14:` / `w15:`. That handles a canonical `<w:sdtPr>` wrapper
 *    whose children inherit a non-`w14` / non-`w15` prefix from the
 *    source's document root.
 */
function normalizeWordPrefix(raw: string, source: XmlElement): string {
  let out = raw;
  // Step 1: wrapper prefix → canonical w.
  const name = source.name ?? "";
  const colonIdx = name.indexOf(":");
  if (colonIdx > 0) {
    const sourcePrefix = name.slice(0, colonIdx);
    if (sourcePrefix !== "w") {
      out = rewritePrefix(out, sourcePrefix, "w");
    }
  }
  // Step 2: child elements that live in known sibling namespaces.
  // - W_LOCAL_NAMES handles inherited alt-prefix SDT property children
  //   sitting inside a canonical `<w:sdtPr>` wrapper (case the pass-19
  //   wrapper-only fix missed).
  // - W14 / W15 sets handle the wider w14:checkbox / w15:repeatingSection
  //   children that don't share local names with run / paragraph
  //   formatting.
  out = normalizeChildrenForLocalNames(out, W_LOCAL_NAMES, "w");
  out = normalizeChildrenForLocalNames(out, W14_LOCAL_NAMES, "w14");
  out = normalizeChildrenForLocalNames(out, W15_LOCAL_NAMES, "w15");
  return out;
}

function rewritePrefix(raw: string, from: string, to: string): string {
  const escaped = from.replaceAll(/[$()*+./?[\\\]^{|}]/gu, "\\$&");
  const tagOpen = new RegExp(`<${escaped}:`, "gu");
  const tagClose = new RegExp(`</${escaped}:`, "gu");
  const attr = new RegExp(`(\\s)${escaped}:`, "gu");
  return raw
    .replaceAll(tagOpen, `<${to}:`)
    .replaceAll(tagClose, `</${to}:`)
    .replaceAll(attr, `$1${to}:`);
}

function normalizeChildrenForLocalNames(
  raw: string,
  localNames: ReadonlySet<string>,
  canonical: string,
): string {
  let out = raw;
  // For each local name, find any `<prefix:localName` and `</prefix:localName>`
  // whose prefix is NOT already canonical, and swap that prefix to canonical.
  // The `\b` after the local name keeps `<prefix:checkboxFoo>` from matching.
  for (const local of localNames) {
    const escapedLocal = local.replaceAll(/[$()*+./?[\\\]^{|}]/gu, "\\$&");
    const opener = new RegExp(`<(\\w+):${escapedLocal}\\b`, "gu");
    const closer = new RegExp(`</(\\w+):${escapedLocal}\\b`, "gu");
    out = out.replaceAll(opener, (_m, prefix: string) =>
      prefix === canonical ? `<${prefix}:${local}` : `<${canonical}:${local}`,
    );
    out = out.replaceAll(closer, (_m, prefix: string) =>
      prefix === canonical ? `</${prefix}:${local}` : `</${canonical}:${local}`,
    );
    // Normalize attribute-name prefixes on the targeted element. Regex-only
    // approaches don't track quote state, so `[^>]{0,200}\s(\w+):` would
    // greedily consume past a quoted attribute value boundary and rewrite
    // a prefix-shaped token sitting *inside* the value
    // (e.g. `<w:tag w:val="foo od:repeat=x0"/>` → `w:val="foo w:repeat=x0"`).
    // Walk the open tag instead, tracking quote state so the rewrite only
    // fires at real attribute-name positions.
    out = rewriteOpenTagAttrPrefixes(out, canonical, local);
  }
  return out;
}

/**
 * Walk every `<canonical:local …>` open tag in `raw` and rewrite any
 * attribute-name prefix that is not already `canonical` (and not `xmlns`) to
 * `canonical`. Quoted attribute values are skipped so prefix-shaped substrings
 * inside a value (OpenDoPE payloads in `w:tag w:val="…"`, namespace URIs in
 * `xmlns:foo="…"`, etc.) are left untouched.
 */
function rewriteOpenTagAttrPrefixes(raw: string, canonical: string, local: string): string {
  const opener = `<${canonical}:${local}`;
  const pieces: string[] = [];
  let cursor = 0;
  while (cursor < raw.length) {
    const tagStart = raw.indexOf(opener, cursor);
    if (tagStart === -1) {
      pieces.push(raw.slice(cursor));
      break;
    }
    // Require a word boundary after the opener so `<w:checkbox` does not
    // also match `<w:checkboxFoo`.
    const afterOpener = raw.codePointAt(tagStart + opener.length);
    const isBoundary =
      afterOpener === undefined ||
      !(
        (
          (afterOpener >= 0x30 && afterOpener <= 0x39) || // 0-9
          (afterOpener >= 0x41 && afterOpener <= 0x5a) || // A-Z
          (afterOpener >= 0x61 && afterOpener <= 0x7a) || // a-z
          afterOpener === 0x5f
        ) // _
      );
    if (!isBoundary) {
      pieces.push(raw.slice(cursor, tagStart + opener.length));
      cursor = tagStart + opener.length;
      continue;
    }
    pieces.push(raw.slice(cursor, tagStart + opener.length));
    cursor = tagStart + opener.length;
    // Walk the open tag to its closing `>`, tracking quote state. Attribute
    // names always sit outside quotes; prefix-shaped tokens inside quotes are
    // attribute values and must be preserved verbatim.
    let quote: '"' | "'" | null = null;
    let atAttrNameBoundary = true;
    while (cursor < raw.length) {
      const ch = raw.charAt(cursor);
      if (quote) {
        pieces.push(ch);
        cursor += 1;
        if (ch === quote) {
          quote = null;
          // Whitespace between attributes must precede the next attribute
          // name; flagging the boundary keeps the rewrite anchored to real
          // attribute-name positions.
          atAttrNameBoundary = false;
        }
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        pieces.push(ch);
        cursor += 1;
        continue;
      }
      if (ch === ">") {
        pieces.push(ch);
        cursor += 1;
        break;
      }
      if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
        pieces.push(ch);
        cursor += 1;
        atAttrNameBoundary = true;
        continue;
      }
      if (atAttrNameBoundary) {
        // Read the next token until `=`, whitespace, `/`, or `>`. If it
        // matches `prefix:localAttr` with a non-canonical, non-xmlns prefix,
        // rewrite the prefix.
        const tokenStart = cursor;
        while (cursor < raw.length) {
          const c = raw.charAt(cursor);
          if (
            c === "=" ||
            c === " " ||
            c === "\t" ||
            c === "\n" ||
            c === "\r" ||
            c === "/" ||
            c === ">"
          ) {
            break;
          }
          cursor += 1;
        }
        const token = raw.slice(tokenStart, cursor);
        const colonIdx = token.indexOf(":");
        if (colonIdx > 0) {
          const prefix = token.slice(0, colonIdx);
          const localAttr = token.slice(colonIdx + 1);
          if (
            prefix !== canonical &&
            prefix !== "xmlns" &&
            /^\w+$/u.test(prefix) &&
            localAttr.length > 0
          ) {
            pieces.push(`${canonical}:${localAttr}`);
          } else {
            pieces.push(token);
          }
        } else {
          pieces.push(token);
        }
        atAttrNameBoundary = false;
        continue;
      }
      // Outside quotes, not at an attribute-name boundary — must be the `=`
      // or stray characters before a quoted value. Copy verbatim.
      pieces.push(ch);
      cursor += 1;
    }
  }
  return pieces.join("");
}

/**
 * The `w:sdtPr` children that state which kind of control this is.
 *
 * They are the choice that closes `CT_SdtPr`, so at most one may appear, and
 * every one of them is kept as bytes rather than rebuilt: `w:date` carries a
 * calendar and a locale, `w:docPartObj` a gallery and a category, `w:text` a
 * `@w:multiLine`, and `w14:checkbox` its two glyph elements — none of which
 * {@link SdtProperties} has a field for. The handler reads the projection
 * folio does model off the element and hands the element itself back, so the
 * kind survives an edit whole.
 *
 * `checkbox` is `w14:` and the Transitional graph does not declare it, so it
 * is listed separately: the dispatcher reaches it through `undeclared` and it
 * cannot be bound to the generated child set.
 */
const CONTROL_KIND_CHILDREN = [
  "equation",
  "comboBox",
  "date",
  "docPartObj",
  "docPartList",
  "dropDownList",
  "picture",
  "richText",
  "text",
  "citation",
  "group",
  "bibliography",
] as const satisfies readonly DeclaredChild<"content-control-properties">[];

const CHECKBOX_CHILD = "checkbox";

/**
 * The local name of a captured fragment's root element.
 *
 * The fragments in the sink are folio's own output — `captureVerbatimXml`
 * serialises one element and materialises its namespace bindings — so the
 * start tag is the first thing in the string and reading its name off the
 * front is exact rather than a guess about arbitrary markup.
 */
const capturedRootName = (xml: string): string | undefined =>
  /^<(?:[\w.-]+:)?(?<local>[\w.-]+)/u.exec(xml)?.groups?.["local"];

/**
 * One preserved `w:sdtPr` child by local name, as the source wrote it.
 *
 * The name is resolved, never a spelling: a producer may bind the w14 URI to
 * any prefix, and the local name is what tells `w14:checkbox` from a run's
 * `w:color`, because `CT_SdtPr` declares neither name twice.
 */
export const preservedSdtChild = (
  preserved: PreservedMarkup | undefined,
  localName: string,
): string | undefined =>
  preserved?.children?.find(({ xml }) => capturedRootName(xml) === localName)?.xml;

/** The same set without that child, or `undefined` when nothing is left. */
export const withoutPreservedSdtChild = (
  preserved: PreservedMarkup | undefined,
  localName: string,
): PreservedMarkup | undefined => {
  const children = (preserved?.children ?? []).filter(
    ({ xml }) => capturedRootName(xml) !== localName,
  );
  return children.length === 0 ? undefined : { children };
};

/**
 * Whether one preserved child is the element that states the control kind.
 *
 * The writer synthesises a kind marker only when the sink holds none, so a
 * control folio parsed keeps the one its author wrote and a control built in
 * code still gets one; and it is this child, not the whole set, that the
 * modelled interactive state is written back into.
 */
export const statesControlKind = (xml: string): boolean => {
  const root = capturedRootName(xml);
  return (
    root !== undefined &&
    (root === CHECKBOX_CHILD || CONTROL_KIND_CHILDREN.some((name) => name === root))
  );
};

/**
 * Parse `<w:sdtPr>` (and optional `<w:sdtEndPr>`) into {@link SdtProperties}.
 *
 * Modelled fields drive addressing and template tooling. Everything else —
 * the control-kind element, `w:dataBinding`, `w:label`, `w:tabIndex`,
 * `w:temporary`, `w:rPr`, and every extension-namespace child — lands in
 * {@link SdtProperties.preserved} at its schema ordinal. A modelled child
 * whose value the reader refuses lands there too: `keptUnless` decides by the
 * outcome, because a map keyed by name cannot list the values a reader will
 * not admit.
 */
export function parseSdtProperties(
  sdtPr: XmlElement | null | undefined,
  sdtEndPr?: XmlElement | null | undefined,
): SdtProperties {
  const props: SdtProperties = { sdtType: "richText" };

  if (sdtPr) {
    /** Read a projection off a control-kind element, and keep the element. */
    const kind =
      (read: (element: XmlElement) => void) =>
      (element: XmlElement): typeof CAPTURE => {
        read(element);
        return CAPTURE;
      };
    const handlers: ChildHandlers<"content-control-properties"> = {
      // The control's placeholder run properties. folio models no run
      // formatting for a control, so the element is kept whole; an empty
      // `<w:rPr/>` is kept as the empty element the source wrote.
      rPr: CAPTURE,
      alias: (element) => {
        const value = getAttributeAnyPrefix(element, "val");
        if (value !== null) {
          props.alias = value;
        }
        return keptUnless(value !== null);
      },
      tag: (element) => {
        const value = getAttributeAnyPrefix(element, "val");
        if (value !== null) {
          props.tag = value;
        }
        return keptUnless(value !== null);
      },
      id: (element) => {
        const raw = getAttributeAnyPrefix(element, "val");
        const value = raw === null ? Number.NaN : Number.parseInt(raw, 10);
        if (!Number.isNaN(value)) {
          props.id = value;
        }
        return keptUnless(!Number.isNaN(value));
      },
      lock: (element) => {
        // `ST_Lock` has four values and the reader admits exactly those. A
        // fifth one used to become `unlocked`, which is a silent change of
        // meaning rather than a loss; now the element keeps its own bytes.
        const value = narrowEnum(getAttributeAnyPrefix(element, "val"), SdtLockSchema);
        if (value !== undefined) {
          props.lock = value;
        }
        return keptUnless(value !== undefined);
      },
      placeholder: (element) => {
        // OOXML §17.5.2.27: the placeholder reference is `w:val` on the
        // nested `<w:docPart>`, not a `<w:val>` child of `w:placeholder`.
        const value = getAttributeAnyPrefix(findChild(element, "w", "docPart"), "val");
        if (value !== null) {
          props.placeholder = value;
        }
        return keptUnless(value !== null);
      },
      // `w:temporary` says the control removes itself once its content is
      // edited. Nothing in folio acts on it, so it travels as bytes.
      temporary: CAPTURE,
      showingPlcHdr: (element) => {
        // OOXML OnOff: a present element with no `w:val` means true, and
        // `0`/`false`/`off` means false. The field is tri-state — absent,
        // on, off — and the writer spells all three, so an explicit
        // negation is not silently turned into an absence.
        props.showingPlaceholder = parseBooleanElement(element, "w");
      },
      // An XPath into the custom-XML store plus the store's id. Keeping it
      // as bytes is honest today; modelling it is a separate change with a
      // consumer of its own.
      dataBinding: CAPTURE,
      label: CAPTURE,
      tabIndex: CAPTURE,
      equation: kind(() => {}),
      comboBox: kind((element) => {
        props.sdtType = "comboBox";
        props.listItems = parseListItems(element);
        readLastValue(element, props);
      }),
      date: kind((element) => {
        props.sdtType = "date";
        // The display format is the child `<w:dateFormat w:val="…"/>`; the
        // bound value is the parent's `@w:fullDate`.
        const format = getAttributeAnyPrefix(findChild(element, "w", "dateFormat"), "val");
        if (format !== null) {
          props.dateFormat = format;
        }
        const fullDate = getAttributeAnyPrefix(element, "fullDate");
        if (fullDate !== null) {
          props.dateValueISO = fullDate;
        }
      }),
      docPartObj: kind(() => {
        props.sdtType = "buildingBlockGallery";
      }),
      docPartList: kind(() => {
        props.sdtType = "buildingBlockGallery";
      }),
      dropDownList: kind((element) => {
        props.sdtType = "dropdown";
        props.listItems = parseListItems(element);
        readLastValue(element, props);
      }),
      picture: kind(() => {
        props.sdtType = "picture";
      }),
      richText: kind(() => {}),
      text: kind(() => {
        props.sdtType = "plainText";
      }),
      citation: kind(() => {}),
      group: kind(() => {
        props.sdtType = "group";
      }),
      bibliography: kind(() => {}),
    };

    const preserved = dispatchChildren({
      element: sdtPr,
      container: "content-control-properties",
      handlers,
      capturePosition: sequencePositions("content-control-properties", sdtPr),
      undeclared: {
        // The checkbox marker is `w14:checkbox`, which the Transitional
        // content model does not declare, so the handler map cannot be total
        // over it and the sink would take it by default. Naming it here is
        // the claim that folio reads a checked state off it; the element
        // itself still goes to the sink, glyph elements and all.
        [CHECKBOX_CHILD]: kind((element) => {
          props.sdtType = "checkbox";
          // OnOff again: `1`/`true`/`on`, or a bare `<w14:checked/>`, is
          // checked. A `w14:checkbox` with no state at all is unchecked.
          const state = findChild(element, "w14", "checked") ?? findChild(element, "w", "checked");
          props.checked = state === null ? false : parseBooleanElement(state, "w14");
        }),
      },
    });
    if (preserved !== undefined) {
      props.preserved = preserved;
    }
  }

  if (sdtEndPr) {
    props.rawEndPropertiesXml = normalizeWordPrefix(captureVerbatimXml(sdtEndPr), sdtEndPr);
  }

  return props;
}

/**
 * The value a producer last chose, from the list element's `@w:lastValue`.
 *
 * `""` is a value a producer can author, so presence is the test rather than
 * truthiness: "never selected", "cleared" and "selected" are three states and
 * the body's display text is evidence for none of them.
 */
function readLastValue(element: XmlElement, props: SdtProperties): void {
  const lastValue = getAttributeAnyPrefix(element, "lastValue");
  if (lastValue !== null) {
    props.dropdownLastValue = lastValue;
  }
}
