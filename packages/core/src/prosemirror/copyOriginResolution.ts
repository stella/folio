import { panic } from "better-result";
import type { Mark, Node as PMNode } from "prosemirror-model";

import {
  acceptCopyOriginAncestors,
  type CopyOriginAncestorList,
  rejectCopyOriginAncestors,
} from "@stll/docx-core/ops";

import { expectBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import type { RevisionSelection } from "./keptRangeBoundaries";
import { expectMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";
import { expectRangeAnchorAttrs } from "./rangeAnchorAttrs";
import type { TrackedRevisionAncestor } from "./schema/marks";
import { trackedRevisionPathOf } from "./trackedRevisionPath";

const isTrackedChangeMark = (mark: Mark): boolean =>
  mark.type.name === "insertion" || mark.type.name === "deletion";

const isAddedLayer = (layer: TrackedRevisionAncestor): boolean =>
  layer.type === "insertion" || layer.type === "moveTo";

const sameRevisionLayer = (
  left: TrackedRevisionAncestor,
  right: TrackedRevisionAncestor,
): boolean => left.type === right.type && left.revisionId === right.revisionId;

const sameOrigins = (
  left: CopyOriginAncestorList | undefined,
  right: CopyOriginAncestorList | undefined,
): boolean =>
  left === right ||
  (left !== undefined &&
    right !== undefined &&
    left.length === right.length &&
    left.every((ordinal, index) => ordinal === right[index]));

const trackedRevisionPathForBoundary = (
  node: PMNode,
): readonly TrackedRevisionAncestor[] => {
  const paths = node.marks
    .filter(isTrackedChangeMark)
    .map(trackedRevisionPathOf);
  const path = paths
    .toSorted((left, right) => right.length - left.length)
    .at(0);
  if (path === undefined) return [];
  if (
    paths.some(
      (candidate) =>
        candidate.length > path.length ||
        candidate.some(
          (layer, index) =>
            path[index] === undefined || !sameRevisionLayer(layer, path[index]),
        ),
    )
  ) {
    return panic(
      "A copied boundary carries conflicting tracked revision paths.",
    );
  }
  return path;
};

type CopyOriginBoundary = {
  kind: "bookmark" | "move";
  origins: CopyOriginAncestorList | undefined;
};

const copyOriginBoundaryOf = (node: PMNode): CopyOriginBoundary | undefined => {
  if (
    node.type.name === "bookmarkBoundary" ||
    node.type.name === "blockBookmarkBoundary"
  ) {
    return {
      kind: "bookmark",
      origins: expectBookmarkBoundaryAttrs(node).copyOriginAncestors,
    };
  }
  if (
    node.type.name === "moveRangeBoundary" ||
    node.type.name === "blockMoveRangeBoundary"
  ) {
    return {
      kind: "move",
      origins: expectMoveRangeBoundaryAttrs(node).copyOriginAncestors,
    };
  }
  return undefined;
};

type WithCopyOriginAncestorsOptions = {
  node: PMNode;
  boundary: CopyOriginBoundary;
  origins: CopyOriginAncestorList | undefined;
};

const withCopyOriginAncestors = ({
  node,
  boundary,
  origins,
}: WithCopyOriginAncestorsOptions): PMNode => {
  if (boundary.kind === "bookmark") {
    return node.type.create(
      {
        ...node.attrs,
        copyOriginAncestors: origins ?? null,
      },
      node.content,
      node.marks,
    );
  }
  const marker = expectMoveRangeBoundaryAttrs(node);
  const { copyOriginAncestors: _previousOrigins, ...withoutOrigins } = marker;
  return node.type.create(
    {
      ...node.attrs,
      marker:
        origins === undefined
          ? withoutOrigins
          : { ...withoutOrigins, copyOriginAncestors: origins },
    },
    node.content,
    node.marks,
  );
};

type WithRangeAnchorCopyOriginsOptions = {
  node: PMNode;
  startOrigins: CopyOriginAncestorList | undefined;
  endOrigins: CopyOriginAncestorList | undefined;
};

const withRangeAnchorCopyOrigins = ({
  node,
  startOrigins,
  endOrigins,
}: WithRangeAnchorCopyOriginsOptions): PMNode => {
  const { start, end } = expectRangeAnchorAttrs(node);
  if (start.type === "commentRangeStart" || end.type === "commentRangeEnd") {
    return node;
  }
  return node.type.create(
    {
      ...node.attrs,
      start: markerWithCopyOrigin(start, startOrigins),
      end: markerWithCopyOrigin(end, endOrigins),
    },
    node.content,
    node.marks,
  );
};

const markerWithCopyOrigin = <
  Marker extends { copyOriginAncestors?: CopyOriginAncestorList },
>(
  marker: Marker,
  origins: CopyOriginAncestorList | undefined,
) => {
  const { copyOriginAncestors: _previousOrigins, ...withoutOrigins } = marker;
  return origins === undefined
    ? withoutOrigins
    : { ...withoutOrigins, copyOriginAncestors: origins };
};

const moveRangeNodeFor = (node: PMNode, marker: unknown): PMNode => {
  const type = node.type.schema.nodes["moveRangeBoundary"];
  if (type === undefined)
    return panic("A move range boundary node is missing.");
  return type.create({ marker }, null, node.marks);
};

type CopyOriginAfterResolutionOptions = {
  origins: CopyOriginAncestorList | undefined;
  path: readonly TrackedRevisionAncestor[];
  selection: RevisionSelection;
};

const copyOriginAfterResolution = ({
  origins,
  path,
  selection,
}: CopyOriginAfterResolutionOptions):
  | { type: "drop" }
  | { type: "retain"; copyOriginAncestors?: CopyOriginAncestorList } => {
  if (origins === undefined) return { type: "retain" };
  const addedLayers = path.filter(isAddedLayer);
  if (
    addedLayers.length === 0 ||
    origins.some((ordinal) => ordinal >= addedLayers.length)
  ) {
    return panic("A copied boundary has no matching added-revision ancestry.");
  }

  if (selection.mode === "accept") {
    let copyOriginAncestors: CopyOriginAncestorList | undefined = origins;
    for (let index = path.length - 1; index >= 0; index -= 1) {
      const layer = path[index];
      if (
        layer === undefined ||
        !isAddedLayer(layer) ||
        !selection.resolves(layer.revisionId)
      )
        continue;
      const acceptedOrdinal = path
        .slice(index + 1)
        .filter(
          (descendant) =>
            isAddedLayer(descendant) &&
            !selection.resolves(descendant.revisionId),
        ).length;
      copyOriginAncestors = acceptCopyOriginAncestors(
        copyOriginAncestors,
        acceptedOrdinal,
      );
    }
    if (sameOrigins(copyOriginAncestors, origins)) {
      return { type: "retain", copyOriginAncestors: origins };
    }
    return copyOriginAncestors === undefined
      ? { type: "retain" }
      : { type: "retain", copyOriginAncestors };
  }

  const selectedOwnerIndex = path.findIndex(
    (layer) => isAddedLayer(layer) && selection.resolves(layer.revisionId),
  );
  if (selectedOwnerIndex < 0) {
    return { type: "retain", copyOriginAncestors: origins };
  }
  const removedAddedAncestors = path
    .slice(selectedOwnerIndex)
    .filter(isAddedLayer).length;
  return rejectCopyOriginAncestors(origins, removedAddedAncestors);
};

/** Resolve copied-boundary ownership using the same nearest-first ordinals as DOCX ops. */
export const copyOriginBoundaryAfterResolution = (
  node: PMNode,
  selection: RevisionSelection,
): PMNode | null => {
  if (node.type.name === "rangeAnchor") {
    const { start, end } = expectRangeAnchorAttrs(node);
    if (start.type === "commentRangeStart" || end.type === "commentRangeEnd") {
      return node;
    }
    const path = trackedRevisionPathForBoundary(node);
    const startResult = copyOriginAfterResolution({
      origins: start.copyOriginAncestors,
      path,
      selection,
    });
    const endResult = copyOriginAfterResolution({
      origins: end.copyOriginAncestors,
      path,
      selection,
    });
    if (startResult.type === "drop") {
      if (endResult.type === "drop") return null;
      return moveRangeNodeFor(
        node,
        markerWithCopyOrigin(end, endResult.copyOriginAncestors),
      );
    }
    if (endResult.type === "drop") {
      return moveRangeNodeFor(
        node,
        markerWithCopyOrigin(start, startResult.copyOriginAncestors),
      );
    }
    if (
      sameOrigins(start.copyOriginAncestors, startResult.copyOriginAncestors) &&
      sameOrigins(end.copyOriginAncestors, endResult.copyOriginAncestors)
    ) {
      return node;
    }
    return withRangeAnchorCopyOrigins({
      node,
      startOrigins: startResult.copyOriginAncestors,
      endOrigins: endResult.copyOriginAncestors,
    });
  }
  const boundary = copyOriginBoundaryOf(node);
  if (boundary === undefined || boundary.origins === undefined) return node;

  const path = trackedRevisionPathForBoundary(node);
  const result = copyOriginAfterResolution({
    origins: boundary.origins,
    path,
    selection,
  });
  if (
    result.type === "retain" &&
    sameOrigins(result.copyOriginAncestors, boundary.origins)
  ) {
    return node;
  }
  return result.type === "drop"
    ? null
    : withCopyOriginAncestors({
        node,
        boundary,
        origins: result.copyOriginAncestors,
      });
};

export const copyOriginBoundaryAfterRemovingAddedAncestors = (
  node: PMNode,
): PMNode | null => {
  if (node.type.name === "rangeAnchor") {
    const { start, end } = expectRangeAnchorAttrs(node);
    if (start.type === "commentRangeStart" || end.type === "commentRangeEnd") {
      return node;
    }
    const addedAncestors =
      trackedRevisionPathForBoundary(node).filter(isAddedLayer).length;
    const remove = (origins: CopyOriginAncestorList | undefined) => {
      if (origins?.some((ordinal) => ordinal >= addedAncestors)) {
        return panic("A copied boundary has no matching removed ancestry.");
      }
      return rejectCopyOriginAncestors(origins, addedAncestors);
    };
    const startResult = remove(start.copyOriginAncestors);
    const endResult = remove(end.copyOriginAncestors);
    if (startResult.type === "drop") {
      if (endResult.type === "drop") return null;
      return moveRangeNodeFor(
        node,
        markerWithCopyOrigin(end, endResult.copyOriginAncestors),
      );
    }
    if (endResult.type === "drop") {
      return moveRangeNodeFor(
        node,
        markerWithCopyOrigin(start, startResult.copyOriginAncestors),
      );
    }
    if (
      sameOrigins(start.copyOriginAncestors, startResult.copyOriginAncestors) &&
      sameOrigins(end.copyOriginAncestors, endResult.copyOriginAncestors)
    ) {
      return node;
    }
    return withRangeAnchorCopyOrigins({
      node,
      startOrigins: startResult.copyOriginAncestors,
      endOrigins: endResult.copyOriginAncestors,
    });
  }
  const boundary = copyOriginBoundaryOf(node);
  if (boundary === undefined || boundary.origins === undefined) return node;
  const removedAddedAncestors =
    trackedRevisionPathForBoundary(node).filter(isAddedLayer).length;
  if (boundary.origins.some((ordinal) => ordinal >= removedAddedAncestors)) {
    return panic("A copied boundary has no matching removed ancestry.");
  }
  const result = rejectCopyOriginAncestors(
    boundary.origins,
    removedAddedAncestors,
  );
  if (
    result.type === "retain" &&
    sameOrigins(result.copyOriginAncestors, boundary.origins)
  ) {
    return node;
  }
  return result.type === "drop"
    ? null
    : withCopyOriginAncestors({
        node,
        boundary,
        origins: result.copyOriginAncestors,
      });
};

type CopyOriginBoundaryAttrsAfterResolutionOptions = {
  source: PMNode;
  current: PMNode;
  baseAttrs: PMNode["attrs"];
  selection: RevisionSelection;
};

export const copyOriginBoundaryAttrsAfterResolution = ({
  source,
  current,
  baseAttrs,
  selection,
}: CopyOriginBoundaryAttrsAfterResolutionOptions): PMNode["attrs"] | null => {
  const resolved = copyOriginBoundaryAfterResolution(source, selection);
  if (resolved === null) return null;
  const boundary = copyOriginBoundaryOf(source);
  if (boundary === undefined) {
    if (
      source.type.name === "rangeAnchor" &&
      resolved.type.name === "rangeAnchor" &&
      resolved !== source
    ) {
      return {
        ...baseAttrs,
        start: resolved.attrs["start"],
        end: resolved.attrs["end"],
      };
    }
    if (
      source.type.name === "rangeAnchor" &&
      resolved.type.name !== "rangeAnchor"
    ) {
      return panic(
        "A resolved point range changed shape outside subtree removal.",
      );
    }
    return baseAttrs;
  }
  if (boundary.origins === undefined) {
    return baseAttrs;
  }
  if (boundary.kind === "bookmark") {
    return {
      ...baseAttrs,
      copyOriginAncestors: resolved.attrs["copyOriginAncestors"],
    };
  }
  const baseNode = current.type.create(
    baseAttrs,
    current.content,
    current.marks,
  );
  const currentMarker = expectMoveRangeBoundaryAttrs(baseNode);
  const resolvedMarker = expectMoveRangeBoundaryAttrs(resolved);
  const { copyOriginAncestors: _currentOrigins, ...markerWithoutOrigins } =
    currentMarker;
  return {
    ...baseAttrs,
    marker:
      resolvedMarker.copyOriginAncestors === undefined
        ? markerWithoutOrigins
        : {
            ...markerWithoutOrigins,
            copyOriginAncestors: resolvedMarker.copyOriginAncestors,
          },
  };
};
