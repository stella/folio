import type { Node as PMNode } from "prosemirror-model";

import type { ShapeAlternateContent } from "../types/content";
import { canonicalJson } from "../utils/canonicalJson";
import type { AlternateContentAttrs } from "./schema/nodes";

const ALTERNATE_CONTENT_ATTR = "_docxAlternateContent";

/**
 * What the node says about the shape: its attributes and, for a text box, its
 * content. Marks are left out; they carry the run's properties, which sit
 * outside `mc:AlternateContent`.
 *
 * The projection from the model is lossy, so the node is compared with itself
 * as it was projected rather than the rebuilt shape with the parsed one: a
 * round trip that changed nothing would otherwise read as an edit.
 */
const nodeFingerprint = (node: PMNode): string => {
  const { [ALTERNATE_CONTENT_ATTR]: _captured, ...attrs } = node.attrs;
  return canonicalJson({ attrs, content: node.content.toJSON() });
};

const nodeAttributesFingerprint = (node: PMNode): string => {
  const { [ALTERNATE_CONTENT_ATTR]: _captured, ...attrs } = node.attrs;
  return canonicalJson(attrs);
};

/** Attach a captured `mc:AlternateContent` to a freshly projected node. */
export const withAlternateContent = (
  node: PMNode,
  xml: string | undefined,
  source?: ShapeAlternateContent,
): PMNode => {
  if (xml === undefined) {
    return node;
  }
  const alternateContent: AlternateContentAttrs = {
    xml,
    fingerprint: nodeFingerprint(node),
    ...(source === undefined
      ? {}
      : {
          attributesFingerprint: nodeAttributesFingerprint(node),
          shapeFingerprint: source.verbatimFingerprint,
          structureFingerprint: source.verbatimStructureFingerprint,
        }),
  };
  return node.type.create(
    { ...node.attrs, [ALTERNATE_CONTENT_ATTR]: alternateContent },
    node.content,
    node.marks,
  );
};

export const editedAlternateContent = (
  node: PMNode,
  captured: AlternateContentAttrs | undefined,
): AlternateContentAttrs | undefined =>
  captured?.attributesFingerprint === nodeAttributesFingerprint(node)
    ? captured
    : undefined;

/** The captured element, while the node is still the one it was attached to. */
export const unchangedAlternateContentXml = (
  node: PMNode,
  captured: AlternateContentAttrs | undefined,
): string | undefined =>
  captured !== undefined && captured.fingerprint === nodeFingerprint(node)
    ? captured.xml
    : undefined;
