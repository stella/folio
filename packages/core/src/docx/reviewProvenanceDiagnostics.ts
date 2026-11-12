import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { scanStreamingXmlElements } from "@stll/docx-core/xml";

import { sliceText } from "../utils/sliceText";
import type { ParseContext } from "./parseContext";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { reviewProvenanceIssue } from "./reviewResolutionProvenance";
import { getLocalName, resolveAttributeNamespaceUri } from "./xmlParser";

/** The longest value a warning quotes; the rest is elided. */
const QUOTED_VALUE_LIMIT = 200;

/**
 * Report each review provenance value the reader drops: a version, a field or
 * a value folio cannot read. Reading drops them and keeps the revisions; this
 * is how the host learns that it did.
 */
export const diagnoseReviewProvenance = (
  parts: ReadonlyMap<string, string>,
  context: ParseContext,
): void => {
  for (const [part, xml] of parts) {
    if (!xml.includes(FOLIO_REVIEW_HISTORY_NAMESPACE)) continue;
    const scoped = context.scoped({ part });
    scanStreamingXmlElements(xml, ({ element }) => {
      for (const [name, value] of Object.entries(element.attributes ?? {})) {
        if (
          !name.includes(":") ||
          name.startsWith("xmlns:") ||
          resolveAttributeNamespaceUri(element, name) !==
            FOLIO_REVIEW_HISTORY_NAMESPACE
        )
          continue;
        const encoded = String(value);
        const issue = reviewProvenanceIssue(getLocalName(name), encoded);
        if (issue === undefined) continue;
        scoped.warn({
          code: PARSE_WARNING_CODES.reviewProvenanceDropped,
          ...(element.name === undefined ? {} : { element: element.name }),
          value:
            encoded.length > QUOTED_VALUE_LIMIT
              ? `${sliceText({ text: encoded, units: "codeUnits", end: QUOTED_VALUE_LIMIT })}...`
              : encoded,
          detail:
            issue === "unsupportedVersion"
              ? `${getLocalName(name)}: unsupported version`
              : `${getLocalName(name)}: invalid value`,
        });
      }
    });
  }
};
