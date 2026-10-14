import {
  IMPORTED_ID_COLLISION_KINDS,
  PARSE_WARNING_CODES,
  type ImportedIdCollisionKind,
} from "@stll/docx-core/model";
import { paraIdAttribute, textIdAttribute } from "./paraIdAttribute";
import { panic } from "better-result";
import type { ParseContext } from "./parseContext";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import { REVISION_ELEMENT_NAMES } from "./revisionIdNormalization";
import { scanSelectedStreamingXmlElements } from "./selectedXmlIdentityScan";
import { toTransitionalNamespaceUri } from "./transitionalSpelling";
import { resolveNamespaceUri } from "./xmlNamespaceContext";
import { getLocalName, getNamespaceUri, NAMESPACES, type XmlElement } from "./xmlParser";

const WORD_2010_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const PACKAGE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";

type CollisionRule = {
  kind: ImportedIdCollisionKind;
  scope: "part" | "package";
  sites: readonly { namespace: string; elements: ReadonlySet<string> }[];
  attribute: string;
  attributeNamespace: string | null;
  encoding: "decimal" | "hex" | "string";
};

const paragraphIdentityForKind = (kind: ImportedIdCollisionKind, element: XmlElement) => {
  switch (kind) {
    case IMPORTED_ID_COLLISION_KINDS.paraId:
      return paraIdAttribute(element);
    case IMPORTED_ID_COLLISION_KINDS.textId:
      return textIdAttribute(element);
    default:
      return undefined;
  }
};

/** Definition sites only: paired markers and references never claim an identity. */
export const IMPORTED_ID_COLLISION_REGISTRY = {
  bookmark: {
    kind: IMPORTED_ID_COLLISION_KINDS.bookmark,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["bookmarkStart"]) }],
    attribute: "id",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  comment: {
    kind: IMPORTED_ID_COLLISION_KINDS.comment,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["comment"]) }],
    attribute: "id",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  footnote: {
    kind: IMPORTED_ID_COLLISION_KINDS.footnote,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["footnote"]) }],
    attribute: "id",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  endnote: {
    kind: IMPORTED_ID_COLLISION_KINDS.endnote,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["endnote"]) }],
    attribute: "id",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  revision: {
    kind: IMPORTED_ID_COLLISION_KINDS.revision,
    scope: "package",
    sites: [{ namespace: NAMESPACES.w, elements: REVISION_ELEMENT_NAMES }],
    attribute: "id",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  paraId: {
    kind: IMPORTED_ID_COLLISION_KINDS.paraId,
    scope: "part",
    sites: [
      { namespace: NAMESPACES.w, elements: new Set(["p"]) },
      { namespace: NAMESPACES.w15, elements: new Set(["commentEx"]) },
      { namespace: OOXML_NAMESPACES.w16cex.uri, elements: new Set(["comment"]) },
    ],
    attribute: "paraId",
    attributeNamespace: WORD_2010_NAMESPACE,
    encoding: "hex",
  },
  textId: {
    kind: IMPORTED_ID_COLLISION_KINDS.textId,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["p"]) }],
    attribute: "textId",
    attributeNamespace: WORD_2010_NAMESPACE,
    encoding: "hex",
  },
  numId: {
    kind: IMPORTED_ID_COLLISION_KINDS.numId,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["num"]) }],
    attribute: "numId",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  abstractNumId: {
    kind: IMPORTED_ID_COLLISION_KINDS.abstractNumId,
    scope: "part",
    sites: [{ namespace: NAMESPACES.w, elements: new Set(["abstractNum"]) }],
    attribute: "abstractNumId",
    attributeNamespace: NAMESPACES.w,
    encoding: "decimal",
  },
  drawing: {
    kind: IMPORTED_ID_COLLISION_KINDS.drawing,
    scope: "package",
    sites: [{ namespace: NAMESPACES.wp, elements: new Set(["docPr"]) }],
    attribute: "id",
    attributeNamespace: null,
    encoding: "decimal",
  },
  relationship: {
    kind: IMPORTED_ID_COLLISION_KINDS.relationship,
    scope: "part",
    sites: [{ namespace: PACKAGE_RELATIONSHIP_NAMESPACE, elements: new Set(["Relationship"]) }],
    attribute: "Id",
    attributeNamespace: null,
    encoding: "string",
  },
} as const satisfies Record<ImportedIdCollisionKind, CollisionRule>;

const normalizedIdentity = (
  value: string,
  encoding: CollisionRule["encoding"],
): string | undefined => {
  if (encoding === "string") return value;
  if (encoding === "hex") return /^[0-9a-f]{8}$/iu.test(value) ? value.toUpperCase() : undefined;
  const spelling = value.trim();
  if (!/^[+-]?\d+$/u.test(spelling)) return undefined;
  const digits = spelling.replace(/^[+-]?0*/u, "");
  return digits === "" ? "0" : `${spelling.startsWith("-") ? "-" : ""}${digits}`;
};

/** Reports repeated source definitions without changing the package or choosing a winner. */
export const diagnoseImportedIdCollisions = (
  parts: ReadonlyMap<string, string>,
  context: ParseContext,
): void => {
  const rules = Object.values(IMPORTED_ID_COLLISION_REGISTRY);
  const packageSeen = new Map<string, { part: string; offset: number }>();
  for (const [part, xml] of parts) {
    const partSeen = new Map<string, { part: string; offset: number }>();
    const scoped = context.scoped({ part });
    scanSelectedStreamingXmlElements(xml, ({ element, attributeValueSpans }) => {
      if (element.name === undefined) return;
      const namespace = toTransitionalNamespaceUri(getNamespaceUri(element) ?? "");
      const localName = getLocalName(element.name);
      for (const rule of rules) {
        const kind = rule.kind;
        if (
          !rule.sites.some((site) => namespace === site.namespace && site.elements.has(localName))
        )
          continue;
        const paragraphIdentity = paragraphIdentityForKind(kind, element);
        const seen = rule.scope === "package" ? packageSeen : partSeen;
        for (const [name, rawValue] of Object.entries(element.attributes ?? {})) {
          if (getLocalName(name) !== rule.attribute || rawValue === undefined) continue;
          const colon = name.indexOf(":");
          const attributeNamespace =
            colon < 0
              ? null
              : toTransitionalNamespaceUri(
                  resolveNamespaceUri(element.namespaceScope, name.slice(0, colon)) ?? "",
                );
          if (rule.encoding === "hex") {
            if (String(rawValue) !== paragraphIdentity) continue;
          } else if (attributeNamespace !== rule.attributeNamespace) continue;
          const value = String(rawValue);
          const identity = normalizedIdentity(value, rule.encoding);
          if (identity === undefined) continue;
          const key = `${kind}:${identity}`;
          const span = attributeValueSpans.get(name);
          if (span === undefined) panic("Missing imported identity source span");
          const offset = span.start;
          const first = seen.get(key);
          if (first === undefined) {
            seen.set(key, { part, offset });
            break;
          }
          scoped.warn({
            code: PARSE_WARNING_CODES.duplicateImportedId,
            identityKind: kind,
            element: element.name,
            at: `offset ${offset}`,
            value,
            detail: `${kind} identity repeats the definition at offset ${first.offset} in ${first.part}.`,
          });
          break;
        }
      }
    });
  }
};
