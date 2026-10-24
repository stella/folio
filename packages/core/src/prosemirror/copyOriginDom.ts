import {
  type CopyOriginAncestorList,
  isCopyOriginAncestors,
} from "@stll/docx-core/ops";

import { MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH } from "../reviewHistoryLimits";

export const INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE =
  "data-docx-internal-clipboard";
export const COPY_ORIGIN_ANCESTORS_ATTRIBUTE =
  "data-docx-copy-origin-ancestors";

export const isInternalClipboardDom = (
  dom: HTMLElement,
  expectedToken: string | undefined,
): boolean =>
  expectedToken !== undefined &&
  dom.getAttribute(INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE) === expectedToken;

export const copyOriginAncestorsFromInternalDom = (
  dom: HTMLElement,
  expectedToken: string | undefined,
): CopyOriginAncestorList | undefined | false => {
  if (!isInternalClipboardDom(dom, expectedToken)) return undefined;
  const encoded = dom.getAttribute(COPY_ORIGIN_ANCESTORS_ATTRIBUTE);
  if (encoded === null) return undefined;
  if (encoded.length > MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH) return false;
  try {
    const parsed: unknown = JSON.parse(encoded);
    return isCopyOriginAncestors(parsed) ? parsed : false;
  } catch {
    return false;
  }
};

export const copyOriginDomAttribute = (
  value: CopyOriginAncestorList | undefined,
  internalClipboardToken: string | undefined,
): Record<string, string> =>
  value === undefined || internalClipboardToken === undefined
    ? {}
    : {
        [COPY_ORIGIN_ANCESTORS_ATTRIBUTE]: JSON.stringify(value),
      };
