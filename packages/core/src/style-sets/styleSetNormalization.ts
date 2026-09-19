/**
 * The boundary an external `DocumentStyleSet` crosses to become a document.
 *
 * A style set is portable by design: hosts persist one as JSON, ship it
 * between workspaces, and hand it back months later. That makes it untrusted
 * input on the same footing as a `.docx`, and it can carry the same defects a
 * package can: a style numbering the set never defines, two styles under one
 * id, an initial paragraph style that is not in the set. A set written before
 * folio learned to normalise those still carries them.
 *
 * So it is normalised here, once, on the way in, through the same owners the
 * parser and the extractor use, rather than at a third call site with its own
 * rules. `createEmptyDocument` is the only door, so this is the only place
 * that has to run.
 *
 * `DOCUMENT_STYLE_SET_VERSION` does not move: the shape is unchanged and a v1
 * value is still a v1 value. What changed is that folio repairs one instead of
 * trusting it.
 */

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { panic } from "better-result";

import { getCachedNumberingMap } from "../docx/numberingParser";
import {
  mintDefaultParagraphStyle,
  resolveDefaultParagraphStyle,
} from "../docx/defaultParagraphStyle";
import { normalizeStyleNumberingReferences } from "../docx/numberingReferenceNormalization";
import type { ParseContext } from "../docx/parseContext";
import type { Style } from "../types/document";
import { DOCUMENT_STYLE_SET_VERSION, type DocumentStyleSet } from "./types";

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const DUPLICATE_STYLE_ID_WARNING = PARSE_WARNING_CODES.styleSetDuplicateStyleId;
export const INITIAL_PARAGRAPH_STYLE_WARNING = PARSE_WARNING_CODES.styleSetInitialStyleMissing;
export const UNNUMBERED_STYLE_SET_STYLE_WARNING = PARSE_WARNING_CODES.unnumberedStyle;

const STYLE_SET_PART = "style-set";

/**
 * A normalised copy of the set, leaving the caller's value untouched.
 *
 * The version is the one thing still refused rather than repaired: a value
 * folio has no reader for is programmer misuse, not a document defect, and
 * guessing at its meaning would be worse than saying so.
 */
export const normalizeDocumentStyleSet = (
  styleSet: DocumentStyleSet,
  context?: ParseContext,
): DocumentStyleSet => {
  if (styleSet.version !== DOCUMENT_STYLE_SET_VERSION) {
    return panic(`Unsupported document style set version: ${String(styleSet.version)}`);
  }
  if (styleSet.name.trim().length === 0) {
    return panic("Document style set name cannot be empty");
  }

  const scoped = context?.scoped({ part: STYLE_SET_PART });
  const normalized = structuredClone(styleSet);
  const styles = dropDuplicateStyleIds(normalized.styles.styles, scoped);
  normalized.styles.styles = styles;

  // Same owner as the parser and the extractor: a style naming a `w:num` the
  // set never defines carries the "no numbering" sentinel rather than a
  // reference nothing resolves, which is what used to reach a panic.
  const unnumbered = normalizeStyleNumberingReferences({
    styles,
    numbering: normalized.numbering ? getCachedNumberingMap(normalized.numbering) : undefined,
  });
  for (const styleId of unnumbered.unnumberedStyleIds) {
    scoped?.warn({
      code: UNNUMBERED_STYLE_SET_STYLE_WARNING,
      value: styleId,
      at: `style "${styleId}"`,
    });
  }

  normalized.initialParagraphStyleId = resolveInitialParagraphStyleId({
    styles,
    requested: styleSet.initialParagraphStyleId,
    hasDocDefaults: normalized.styles.docDefaults !== undefined,
    context: scoped,
  });
  return normalized;
};

const dropDuplicateStyleIds = (styles: Style[], context: ParseContext | undefined): Style[] => {
  const seen = new Set<string>();
  const kept: Style[] = [];
  for (const style of styles) {
    if (seen.has(style.styleId)) {
      // A `w:pStyle` names one id, so a repeat is a style nothing can apply;
      // Word resolves a reference to the first definition and so does folio.
      context?.warn({
        code: DUPLICATE_STYLE_ID_WARNING,
        value: style.styleId,
        at: `style "${style.styleId}"`,
      });
      continue;
    }
    seen.add(style.styleId);
    kept.push(style);
  }
  return kept;
};

type ResolveInitialParagraphStyleIdOptions = {
  styles: Style[];
  requested: string;
  /** Whether the set carries `w:docDefaults`, which stay authoritative. */
  hasDocDefaults: boolean;
  context: ParseContext | undefined;
};

const resolveInitialParagraphStyleId = ({
  styles,
  requested,
  hasDocDefaults,
  context,
}: ResolveInitialParagraphStyleIdOptions): string => {
  const named = styles.find((style) => style.styleId === requested && style.type === "paragraph");
  if (named) {
    return requested;
  }
  const resolved = resolveDefaultParagraphStyle(styles);
  const replacement =
    resolved ??
    (() => {
      const minted = mintDefaultParagraphStyle({
        takenStyleIds: new Set(styles.map((style) => style.styleId)),
        hasDocDefaults,
      });
      styles.push(minted);
      return minted;
    })();
  context?.warn({
    code: INITIAL_PARAGRAPH_STYLE_WARNING,
    value: requested,
    at: `style "${replacement.styleId}"`,
  });
  return replacement.styleId;
};
