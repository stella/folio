import type { Node as PMNode } from "prosemirror-model";

import type { TextBoxAnchorAttrs, TextBoxAttrs } from "../../prosemirror/schema/nodes";
import { canonicalJson } from "../../utils/canonicalJson";

type SourceIdentityAttrDisposition = "exact" | "text-box-anchor" | "text-box-group";

/**
 * Text-box group and anchor ids are random per editor load. Their equality
 * relationships carry source structure; their literal values do not. Keeping
 * this map total makes every new text-box attribute an explicit identity
 * decision instead of letting projection drift become invisible.
 */
const TEXT_BOX_SOURCE_IDENTITY_ATTRS = Object.freeze({
  width: "exact",
  height: "exact",
  autoFit: "exact",
  textWrap: "exact",
  textBoxId: "exact",
  fillColor: "exact",
  outlineWidth: "exact",
  outlineColor: "exact",
  outlineStyle: "exact",
  transform: "exact",
  marginTop: "exact",
  marginBottom: "exact",
  marginLeft: "exact",
  marginRight: "exact",
  verticalAlign: "exact",
  displayMode: "exact",
  cssFloat: "exact",
  wrapType: "exact",
  wrapText: "exact",
  distTop: "exact",
  distBottom: "exact",
  distLeft: "exact",
  distRight: "exact",
  position: "exact",
  _docxPlacement: "exact",
  _docxGroupId: "text-box-group",
  _docxAnchorId: "text-box-anchor",
  _docxTrackedChange: "exact",
  _docxInlineSdts: "exact",
} as const satisfies Readonly<Record<keyof TextBoxAttrs, SourceIdentityAttrDisposition>>);

const TEXT_BOX_ANCHOR_SOURCE_IDENTITY_ATTRS = Object.freeze({
  anchorId: "text-box-anchor",
} as const satisfies Readonly<Record<keyof TextBoxAnchorAttrs, SourceIdentityAttrDisposition>>);

type AlphaIdentityState = {
  readonly left: Map<string, number>;
  readonly right: Map<string, number>;
};

type SourceIdentityState = {
  readonly anchors: AlphaIdentityState;
  readonly groups: AlphaIdentityState;
};

const hasOwn = <Value extends object>(value: Value, key: PropertyKey): key is keyof Value =>
  Object.prototype.hasOwnProperty.call(value, key);

const attrDisposition = (node: PMNode, key: string): SourceIdentityAttrDisposition => {
  if (node.type.name === "textBox" && hasOwn(TEXT_BOX_SOURCE_IDENTITY_ATTRS, key)) {
    return TEXT_BOX_SOURCE_IDENTITY_ATTRS[key];
  }
  if (node.type.name === "textBoxAnchor" && hasOwn(TEXT_BOX_ANCHOR_SOURCE_IDENTITY_ATTRS, key)) {
    return TEXT_BOX_ANCHOR_SOURCE_IDENTITY_ATTRS[key];
  }
  return "exact";
};

const alphaIdentity = (value: unknown, identities: Map<string, number>): unknown => {
  if (value === null || value === undefined) return null;
  if (typeof value !== "string") return value;
  const existing = identities.get(value);
  if (existing !== undefined) return existing;
  const identity = identities.size;
  identities.set(value, identity);
  return identity;
};

const validAlphaIdentity = (value: unknown, optional: boolean): boolean =>
  (optional && (value === null || value === undefined)) ||
  (typeof value === "string" && value.length > 0);

const attrsMatch = (left: PMNode, right: PMNode, state: SourceIdentityState): string | null => {
  const attrKeys = new Set([...Object.keys(left.attrs), ...Object.keys(right.attrs)]);
  for (const key of attrKeys) {
    const disposition = attrDisposition(left, key);
    const leftValue = left.attrs[key];
    const rightValue = right.attrs[key];
    if (disposition === "exact") {
      if (canonicalJson(leftValue) !== canonicalJson(rightValue)) return key;
      continue;
    }
    const identityIsOptional = left.type.name === "textBox";
    if (
      !validAlphaIdentity(leftValue, identityIsOptional) ||
      !validAlphaIdentity(rightValue, identityIsOptional)
    ) {
      return key;
    }
    const identities = disposition === "text-box-group" ? state.groups : state.anchors;
    if (alphaIdentity(leftValue, identities.left) !== alphaIdentity(rightValue, identities.right)) {
      return key;
    }
  }
  return null;
};

const firstDifferencePath = (
  left: PMNode,
  right: PMNode,
  state: SourceIdentityState,
  path: string,
): string => {
  if (left.type !== right.type) return `${path}.type`;
  const attrDifference = attrsMatch(left, right, state);
  if (attrDifference !== null) return `${path}.attrs.${attrDifference}`;
  if (left.text !== right.text) return `${path}.text`;
  if (
    canonicalJson(left.marks.map((mark) => mark.toJSON())) !==
    canonicalJson(right.marks.map((mark) => mark.toJSON()))
  ) {
    return `${path}.marks`;
  }
  if (left.childCount !== right.childCount) return `${path}.childCount`;
  for (let index = 0; index < left.childCount; index++) {
    const difference = firstDifferencePath(
      left.child(index),
      right.child(index),
      state,
      `${path}.content[${String(index)}]`,
    );
    if (difference !== "") return difference;
  }
  return "";
};

/**
 * Return the first structural or semantic PM difference. Session-random text
 * box identifiers are alpha-compared, so grouping and anchor links must remain
 * identical even though independently loaded documents have different salts.
 */
export const firstProseMirrorSourceIdentityDifferencePath = (left: PMNode, right: PMNode): string =>
  firstDifferencePath(
    left,
    right,
    {
      anchors: { left: new Map(), right: new Map() },
      groups: { left: new Map(), right: new Map() },
    },
    "doc",
  );
