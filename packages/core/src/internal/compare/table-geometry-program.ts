/**
 * A table's own properties, read and matched across two documents.
 *
 * `w:tblPr`, `w:trPr` and `w:tcPr` are the part of a table no block carries: a
 * paragraph snapshot says which cell it sits in and nothing about the cell's
 * width, span, shading, borders or margins, nor about the row's height or
 * header flag, nor about the table's own width, indent, justification,
 * borders, cell margins, look or style. A comparison that reproduces every
 * word and none of that hands back a table that is not the one it was compared
 * to.
 *
 * Two things happen here. {@link projectTableGeometry} renders those
 * properties as text, so the round-trip self-check sees a lost property the
 * same way it sees a lost paragraph. The preflight builds one immutable,
 * reversible program; its executor writes that program to a caller-owned
 * transaction without consulting the target document.
 *
 * The scope of each is not a hand-kept list. It is exactly the attrs a reject
 * of the matching change element restores, read off the reject patches
 * themselves — so a property that reject can restore is a property the
 * comparison carries and checks, and the three can never drift apart.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { panic, Result } from "better-result";

import { markStructuralChange } from "../../prosemirror/extensions/features/ParagraphChangeTrackerExtension";

import {
  tableCellRejectAttrPatch,
  tableRejectAttrPatch,
  tableRowRejectAttrPatch,
} from "../../prosemirror/commands/propertyChangeScope";
import {
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
} from "../../prosemirror/conversion/fromProseDoc";
import type { TableCellAttrs } from "../../prosemirror/schema/nodes";
import type {
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";
import { canonicalJson } from "../../utils/canonicalJson";
import type { FolioStoryTable } from "../../ai-edits/snapshot";

const ORIGINAL_FORMATTING = "_originalFormatting";

/**
 * The formatting a parser stored verbatim for the save path. It is deliberately
 * outside every projection below: a document saved by an editor materializes
 * style-resolved properties into its own `w:tcPr`, so two packages that render
 * the same table can store different formatting for it, and comparing what was
 * stored would report a difference no redline can or should represent. What is
 * compared is the effective set the properties resolve to.
 */
const withoutOriginalFormatting = (keys: readonly string[]): string[] =>
  keys.filter((key) => key !== ORIGINAL_FORMATTING);

/**
 * Attrs a `w:tblPrChange` reject restores, and therefore the attrs a match
 * writes and the projection reads. `columnWidths` is deliberately outside
 * them: the grid is `w:tblGrid`, not `w:tblPr`, and the editable model has no
 * `w:tblGridChange` to record a change of it against.
 */
const TABLE_SCOPED_ATTRS = withoutOriginalFormatting(Object.keys(tableRejectAttrPatch(undefined)));
const ROW_SCOPED_ATTRS = withoutOriginalFormatting(Object.keys(tableRowRejectAttrPatch(undefined)));
const CELL_SCOPED_ATTRS = withoutOriginalFormatting(
  Object.keys(tableCellRejectAttrPatch(undefined, undefined)),
);

/**
 * `false` and absent are the same thing for every property in scope here: they
 * are all presence flags (`w:tblHeader`, `w:hidden`, `w:noWrap`), and a parser
 * that materializes an absent one as `false` must not read as a difference
 * from one that leaves it unset.
 */
const scopedAttrs = (
  attrs: Record<string, unknown>,
  keys: readonly string[],
): Record<string, unknown> => {
  const scoped: Record<string, unknown> = {};
  for (const key of keys) {
    const value = attrs[key];
    scoped[key] = value === false ? null : (value ?? null);
  }
  return scoped;
};

const TABLE_NODE_ROLE = "table";
const TABLE_ROW_NODE_ROLE = "row";
const TABLE_CELL_NODE_ROLES = new Set(["cell", "header_cell"]);

/** Where a cell sits: which table, which row of it, which cell of that row. */
export type TableCellCoordinate = {
  readonly tableIndex: number;
  readonly rowIndex: number;
  readonly cellIndex: number;
};

/** One base cell and the target cell it was aligned with. */
export type TableGeometryPairing = {
  readonly base: TableCellCoordinate;
  readonly target: TableCellCoordinate;
};

export type TableGeometryPreflightLimits = {
  readonly maxTables: number;
  readonly maxPairings: number;
  readonly maxVisitedNodes: number;
  readonly maxChanges: number;
  readonly maxPayloadUnits: number;
};

/** Hard stops for the internal comparison transport, below the package parser's limits. */
export const DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS = Object.freeze({
  maxTables: 10_000,
  maxPairings: 10_000,
  maxVisitedNodes: 100_000,
  maxChanges: 10_000,
  maxPayloadUnits: 4_194_304,
} as const satisfies TableGeometryPreflightLimits);

type TableGeometryLimit = keyof TableGeometryPreflightLimits;
type TableGeometrySide = "base" | "target";
type TableGeometryScopeName = "table" | "row" | "cell";

export type TableGeometryUnsupportedIssue =
  | {
      readonly reason: "invalid-limit";
      readonly limit: TableGeometryLimit;
      readonly actual: number;
    }
  | {
      readonly reason: "limit-exceeded";
      readonly limit: TableGeometryLimit;
      readonly maximum: number;
      readonly actual: number;
    }
  | {
      readonly reason: "invalid-coordinate";
      readonly side: TableGeometrySide;
      readonly coordinate: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "invalid-table-index" | "duplicate-table-index";
      readonly side: TableGeometrySide;
      readonly tableIndex: number;
    }
  | {
      readonly reason: "invalid-table-position" | "duplicate-table-position";
      readonly position: number;
    }
  | {
      readonly reason: "duplicate-pairing";
      readonly side: TableGeometrySide;
      readonly coordinate: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "conflicting-table-pairing" | "conflicting-row-pairing";
      readonly side: TableGeometrySide;
      readonly coordinate: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "missing-table" | "missing-row" | "missing-cell" | "unexpected-node-role";
      readonly side: TableGeometrySide;
      readonly scope: TableGeometryScopeName;
      readonly coordinate: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "structural-cell-mismatch";
      readonly property: "node-type" | "colspan" | "rowspan";
      readonly base: Readonly<TableCellCoordinate>;
      readonly target: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "non-reconstructable-structure-change";
      readonly scope: "table";
      readonly property: "column-widths";
      readonly base: Readonly<TableCellCoordinate>;
      readonly target: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "non-reconstructable-property-change";
      readonly reconstruction: "accepted" | "rejected";
      readonly scope: TableGeometryScopeName;
      readonly base: Readonly<TableCellCoordinate>;
      readonly target: Readonly<TableCellCoordinate>;
    }
  | {
      readonly reason: "invalid-property-payload";
      readonly scope: TableGeometryScopeName;
      readonly base: Readonly<TableCellCoordinate>;
      readonly target: Readonly<TableCellCoordinate>;
    };

export type TableGeometryExecutionIssue =
  | {
      readonly reason: "invalid-revision-stamp";
    }
  | {
      readonly reason: "missing-live-node" | "stale-live-node" | "invalid-live-node";
      readonly scope: TableGeometryScopeName;
      readonly position: number;
    }
  | {
      readonly reason: "live-payload-limit";
      readonly maximum: number;
      readonly actual: number;
    };

type TableGeometryRevisionStamp = {
  readonly author: string;
  readonly date: string;
  readonly idSeed: number;
};

type TableGeometryRevisionReceipt = {
  readonly scope: TableGeometryScopeName;
  readonly position: number;
  readonly revisionId: number;
  readonly base: Readonly<TableCellCoordinate>;
  readonly target: Readonly<TableCellCoordinate>;
};

export type TableGeometryExecutionReceipt = {
  readonly type: "table-geometry-execution";
  readonly startingRevisionId: number;
  readonly nextRevisionId: number;
  readonly revisions: readonly TableGeometryRevisionReceipt[];
};

type TableGeometryFormattingByScope = {
  readonly table: TableFormatting;
  readonly row: TableRowFormatting;
  readonly cell: TableCellFormatting;
};

type PropertyChangePayloadByScope = {
  readonly table: {
    readonly changeAttr: "tblPrChange";
    readonly changeType: "tablePropertyChange";
    readonly previousFormatting: Readonly<TableFormatting> | undefined;
  };
  readonly row: {
    readonly changeAttr: "trPrChange";
    readonly changeType: "tableRowPropertyChange";
    readonly previousFormatting: Readonly<TableRowFormatting> | undefined;
  };
  readonly cell: {
    readonly changeAttr: "tcPrChange";
    readonly changeType: "tableCellPropertyChange";
    readonly previousFormatting: Readonly<TableCellFormatting> | undefined;
  };
};

type TableGeometryCarrierAssertion<Scope extends TableGeometryScopeName = TableGeometryScopeName> =
  {
    readonly scope: Scope;
    readonly position: number;
    readonly expectedLiveState: string;
    readonly base: Readonly<TableCellCoordinate>;
    readonly target: Readonly<TableCellCoordinate>;
  };

type TableGeometryInstructionFor<Scope extends TableGeometryScopeName> =
  Scope extends TableGeometryScopeName
    ? {
        readonly carrier: TableGeometryCarrierAssertion<Scope>;
        readonly targetAttrs: Readonly<Record<string, unknown>>;
        readonly change: PropertyChangePayloadByScope[Scope];
      }
    : never;

type TableGeometryInstruction = TableGeometryInstructionFor<TableGeometryScopeName>;

const TABLE_GEOMETRY_PROGRAM_BRAND: unique symbol = Symbol("table-geometry-program");

type TableGeometryProgram = {
  readonly [TABLE_GEOMETRY_PROGRAM_BRAND]: true;
  readonly type: "unchanged" | "changes";
  readonly carriers: readonly TableGeometryCarrierAssertion[];
  readonly instructions: readonly TableGeometryInstruction[];
  readonly maxLivePayloadUnits: number;
};

export type TableGeometryPreflightResult =
  | { readonly status: "ready"; readonly program: TableGeometryProgram }
  | { readonly status: "unsupported"; readonly issue: TableGeometryUnsupportedIssue };

export type TableGeometryExecutionResult =
  | { readonly status: "executed"; readonly receipt: TableGeometryExecutionReceipt }
  | { readonly status: "unsupported"; readonly issue: TableGeometryExecutionIssue };

type PropertyScope<Scope extends TableGeometryScopeName> = {
  readonly name: Scope;
  readonly keys: readonly string[];
  /** The complete property set that the node serializes. */
  readonly formattingOf: (node: PMNode) => TableGeometryFormattingByScope[Scope] | undefined;
  /** The semantic portion of that property set; transport source is excluded. */
  readonly semanticFormattingOf: (
    formatting: TableGeometryFormattingByScope[Scope] | undefined,
  ) => unknown;
  /** What rejecting a change element restores from its stored property set. */
  readonly rejectPatch: (
    previousFormatting: TableGeometryFormattingByScope[Scope] | undefined,
    liveFormatting: TableGeometryFormattingByScope[Scope] | undefined,
  ) => Record<string, unknown>;
  readonly changeAttr: PropertyChangePayloadByScope[Scope]["changeAttr"];
  readonly changeType: PropertyChangePayloadByScope[Scope]["changeType"];
  readonly createInstruction: (
    carrier: TableGeometryCarrierAssertion<Scope>,
    targetAttrs: Readonly<Record<string, unknown>>,
    previousFormatting: TableGeometryFormattingByScope[Scope] | undefined,
  ) => TableGeometryInstructionFor<Scope>;
};

/**
 * `Node["attrs"]` is an open record, and the converters below want the node
 * type's own shape. The two span counts are the only members the schema always
 * carries a value for, so naming them is what turns the record into one.
 */
const cellAttrsOf = (node: PMNode): TableCellAttrs => ({
  ...effectiveAttrs(node),
  colspan: typeof node.attrs["colspan"] === "number" ? node.attrs["colspan"] : 1,
  rowspan: typeof node.attrs["rowspan"] === "number" ? node.attrs["rowspan"] : 1,
});

/**
 * The node's attrs with the style cascade's own values cleared, so the
 * converter treats every effective value as one the node states.
 *
 * A change element stores the COMPLETE previous property set and a reject
 * rebuilds the live properties from it alone, so the record has to hold what
 * the node renders with — including a border its table style supplied. The
 * save path wants the opposite (write only what the node states, or the
 * inherited value becomes an override), which is what the resolved companions
 * are for; a record is the one place they get in the way.
 */
const effectiveAttrs = (node: PMNode): Record<string, unknown> => ({
  ...node.attrs,
  _resolvedBackgroundColor: null,
  _resolvedBorders: null,
  _resolvedMargins: null,
  _resolvedCellMargins: null,
});

const tableSemanticFormatting = (formatting: TableFormatting | undefined): unknown => {
  if (!formatting) return null;
  const { sourceXml: _sourceXml, gridSourceXml: _gridSourceXml, ...semantic } = formatting;
  return semantic;
};

const rowSemanticFormatting = (formatting: TableRowFormatting | undefined): unknown => {
  if (!formatting) return null;
  const { sourceXml: _sourceXml, ...semantic } = formatting;
  return semantic;
};

const cellSemanticFormatting = (formatting: TableCellFormatting | undefined): unknown => {
  if (!formatting) return null;
  const {
    sourceXml: _sourceXml,
    gridSpan: _gridSpan,
    vMerge: _verticalMerge,
    ...semantic
  } = formatting;
  return semantic;
};

const TABLE_SCOPE = {
  name: "table",
  keys: TABLE_SCOPED_ATTRS,
  formattingOf: (node) => tableAttrsToFormatting(effectiveAttrs(node)),
  semanticFormattingOf: tableSemanticFormatting,
  rejectPatch: (previousFormatting) => tableRejectAttrPatch(previousFormatting),
  changeAttr: "tblPrChange",
  changeType: "tablePropertyChange",
  createInstruction: (carrier, targetAttrs, previousFormatting) =>
    Object.freeze({
      carrier,
      targetAttrs,
      change: Object.freeze({
        changeAttr: "tblPrChange",
        changeType: "tablePropertyChange",
        previousFormatting,
      }),
    }),
} as const satisfies PropertyScope<"table">;

const ROW_SCOPE = {
  name: "row",
  keys: ROW_SCOPED_ATTRS,
  formattingOf: (node) => tableRowAttrsToFormatting(node.attrs),
  semanticFormattingOf: rowSemanticFormatting,
  rejectPatch: (previousFormatting) => tableRowRejectAttrPatch(previousFormatting),
  changeAttr: "trPrChange",
  changeType: "tableRowPropertyChange",
  createInstruction: (carrier, targetAttrs, previousFormatting) =>
    Object.freeze({
      carrier,
      targetAttrs,
      change: Object.freeze({
        changeAttr: "trPrChange",
        changeType: "tableRowPropertyChange",
        previousFormatting,
      }),
    }),
} as const satisfies PropertyScope<"row">;

const CELL_SCOPE = {
  name: "cell",
  keys: CELL_SCOPED_ATTRS,
  formattingOf: (node) => tableCellAttrsToFormatting(cellAttrsOf(node)),
  semanticFormattingOf: cellSemanticFormatting,
  rejectPatch: (previousFormatting, liveFormatting) =>
    tableCellRejectAttrPatch(previousFormatting, liveFormatting),
  changeAttr: "tcPrChange",
  changeType: "tableCellPropertyChange",
  createInstruction: (carrier, targetAttrs, previousFormatting) =>
    Object.freeze({
      carrier,
      targetAttrs,
      change: Object.freeze({
        changeAttr: "tcPrChange",
        changeType: "tableCellPropertyChange",
        previousFormatting,
      }),
    }),
} as const satisfies PropertyScope<"cell">;

type PayloadCaptureFailure =
  | { readonly type: "invalid" }
  | { readonly type: "limit"; readonly actual: number };

type PayloadCaptureContext = {
  readonly maximum: number;
  used: number;
  readonly active: WeakSet<object>;
  readonly captured: WeakMap<object, unknown>;
};

const consumePayloadUnits = (
  context: PayloadCaptureContext,
  units: number,
): Result<void, PayloadCaptureFailure> => {
  context.used += units;
  return context.used <= context.maximum
    ? Result.ok(undefined)
    : Result.err({ type: "limit", actual: context.used });
};

/** Capture plain semantic payload once; no caller-owned object reaches an executable program. */
const captureImmutablePayload = <Value>(
  value: Value,
  context: PayloadCaptureContext,
  depth = 0,
): Result<Value, PayloadCaptureFailure> => {
  if (depth > 64) return Result.err({ type: "invalid" });
  if (value === null || value === undefined || typeof value === "boolean") {
    return Result.ok(value);
  }
  if (typeof value === "string") {
    const consumed = consumePayloadUnits(context, value.length + 1);
    return consumed.isErr() ? Result.err(consumed.error) : Result.ok(value);
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return Result.err({ type: "invalid" });
    const consumed = consumePayloadUnits(context, 1);
    return consumed.isErr() ? Result.err(consumed.error) : Result.ok(value);
  }
  if (typeof value !== "object") return Result.err({ type: "invalid" });
  const objectValue: object = value;

  const retained = context.captured.get(objectValue);
  if (retained !== undefined) {
    // SAFETY: retained was captured from this exact object in this generic invocation graph.
    return Result.ok(retained as Value);
  }
  if (context.active.has(objectValue)) return Result.err({ type: "invalid" });
  context.active.add(objectValue);

  const descriptors = Result.try({
    try: () => Object.getOwnPropertyDescriptors(objectValue),
    catch: () => ({ type: "invalid" }) as const,
  });
  if (descriptors.isErr()) {
    context.active.delete(objectValue);
    return Result.err(descriptors.error);
  }

  if (Array.isArray(value)) {
    const lengthDescriptor = descriptors.value["length"];
    if (
      !lengthDescriptor ||
      !("value" in lengthDescriptor) ||
      !Number.isSafeInteger(lengthDescriptor.value) ||
      lengthDescriptor.value < 0
    ) {
      context.active.delete(objectValue);
      return Result.err({ type: "invalid" });
    }
    const length: number = lengthDescriptor.value;
    const consumed = consumePayloadUnits(context, length + 1);
    if (consumed.isErr()) {
      context.active.delete(objectValue);
      return Result.err(consumed.error);
    }
    const clone: unknown[] = [];
    for (let index = 0; index < length; index++) {
      const descriptor = descriptors.value[String(index)];
      if (!descriptor || !("value" in descriptor)) {
        context.active.delete(objectValue);
        return Result.err({ type: "invalid" });
      }
      const entry = captureImmutablePayload(descriptor.value, context, depth + 1);
      if (entry.isErr()) {
        context.active.delete(objectValue);
        return Result.err(entry.error);
      }
      clone.push(entry.value);
    }
    const allowedKeys = new Set(["length", ...clone.map((_entry, index) => String(index))]);
    if (Reflect.ownKeys(descriptors.value).some((key) => !allowedKeys.has(String(key)))) {
      context.active.delete(objectValue);
      return Result.err({ type: "invalid" });
    }
    const frozen = Object.freeze(clone);
    context.active.delete(objectValue);
    context.captured.set(objectValue, frozen);
    // SAFETY: the dense array was recursively cloned without changing its value shape.
    return Result.ok(frozen as Value);
  }

  const prototype = Result.try({
    try: () => Object.getPrototypeOf(value),
    catch: () => ({ type: "invalid" }) as const,
  });
  if (prototype.isErr() || (prototype.value !== Object.prototype && prototype.value !== null)) {
    context.active.delete(objectValue);
    return Result.err(prototype.isErr() ? prototype.error : { type: "invalid" });
  }
  const keys = Reflect.ownKeys(descriptors.value);
  if (keys.some((key) => typeof key !== "string")) {
    context.active.delete(objectValue);
    return Result.err({ type: "invalid" });
  }
  const consumed = consumePayloadUnits(context, keys.length + 1);
  if (consumed.isErr()) {
    context.active.delete(objectValue);
    return Result.err(consumed.error);
  }
  const clone: Record<string, unknown> = {};
  for (const key of keys.toSorted()) {
    if (typeof key !== "string") return panic("A symbol passed the table payload key guard");
    const descriptor = descriptors.value[key];
    if (!descriptor || !("value" in descriptor) || descriptor.enumerable !== true) {
      context.active.delete(objectValue);
      return Result.err({ type: "invalid" });
    }
    const entry = captureImmutablePayload(descriptor.value, context, depth + 1);
    if (entry.isErr()) {
      context.active.delete(objectValue);
      return Result.err(entry.error);
    }
    clone[key] = entry.value;
  }
  const frozen = Object.freeze(clone);
  context.active.delete(objectValue);
  context.captured.set(objectValue, frozen);
  // SAFETY: the plain record was recursively cloned without changing its value shape.
  return Result.ok(frozen as Value);
};

const propertyStateValue = <Scope extends TableGeometryScopeName>(
  node: PMNode,
  scope: PropertyScope<Scope>,
): unknown => ({
  attrs: scopedAttrs(node.attrs, scope.keys),
  formatting: scope.semanticFormattingOf(scope.formattingOf(node)),
});

const propertyState = <Scope extends TableGeometryScopeName>(
  node: PMNode,
  scope: PropertyScope<Scope>,
): string => canonicalJson(propertyStateValue(node, scope));

const capturedState = (
  value: unknown,
  context: PayloadCaptureContext,
): Result<string, PayloadCaptureFailure> => {
  const captured = captureImmutablePayload(value, context);
  return captured.isErr() ? Result.err(captured.error) : Result.ok(canonicalJson(captured.value));
};

const liveStructure = (node: PMNode, scope: TableGeometryScopeName): unknown => {
  switch (scope) {
    case "table":
      return { columnWidths: node.attrs["columnWidths"] ?? null };
    case "row":
      return { childCount: node.childCount };
    case "cell":
      return {
        colspan: node.attrs["colspan"] ?? 1,
        rowspan: node.attrs["rowspan"] ?? 1,
      };
    default: {
      const unreachable: never = scope;
      return panic("Unhandled table geometry scope", { scope: unreachable });
    }
  }
};

const capturedLiveState = <Scope extends TableGeometryScopeName>(
  node: PMNode,
  scope: PropertyScope<Scope>,
  context: PayloadCaptureContext,
): Result<string, PayloadCaptureFailure> =>
  capturedState(
    {
      nodeType: node.type.name,
      properties: propertyStateValue(node, scope),
      structure: liveStructure(node, scope.name),
      existingChanges: node.attrs[scope.changeAttr] ?? null,
    },
    context,
  );

const nodeWithAttrs = (
  node: PMNode,
  attrs: Readonly<Record<string, unknown>>,
): Result<PMNode, PayloadCaptureFailure> =>
  Result.try({
    try: () => node.type.create({ ...node.attrs, ...attrs }, node.content, node.marks),
    catch: () => ({ type: "invalid" }) as const,
  });

type PropertyPreflightResult<Scope extends TableGeometryScopeName> =
  | { readonly status: "unchanged" }
  | { readonly status: "ready"; readonly instruction: TableGeometryInstructionFor<Scope> }
  | { readonly status: "unsupported"; readonly issue: TableGeometryUnsupportedIssue };

type PropertyChangeForOptions<Scope extends TableGeometryScopeName> = {
  readonly baseNode: PMNode;
  readonly targetNode: PMNode;
  readonly scope: PropertyScope<Scope>;
  readonly carrier: TableGeometryCarrierAssertion<Scope>;
  readonly payloadContext: PayloadCaptureContext;
};

const payloadFailureIssue = ({
  failure,
  scope,
  base,
  target,
  maximum,
}: {
  failure: PayloadCaptureFailure;
  scope: TableGeometryScopeName;
  base: Readonly<TableCellCoordinate>;
  target: Readonly<TableCellCoordinate>;
  maximum: number;
}): TableGeometryUnsupportedIssue =>
  failure.type === "limit"
    ? {
        reason: "limit-exceeded",
        limit: "maxPayloadUnits",
        maximum,
        actual: failure.actual,
      }
    : { reason: "invalid-property-payload", scope, base, target };

/** Build one complete reversible instruction, or name why it is unsupported. */
const propertyChangeFor = <Scope extends TableGeometryScopeName>({
  baseNode,
  targetNode,
  scope,
  carrier,
  payloadContext,
}: PropertyChangeForOptions<Scope>): PropertyPreflightResult<Scope> => {
  const { base, target } = carrier;
  const baseFormatting = scope.formattingOf(baseNode);
  const targetFormatting = scope.formattingOf(targetNode);
  const baseState = capturedState(propertyStateValue(baseNode, scope), payloadContext);
  if (baseState.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: baseState.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }
  const targetState = capturedState(propertyStateValue(targetNode, scope), payloadContext);
  if (targetState.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: targetState.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }
  if (baseState.value === targetState.value) return { status: "unchanged" };

  const targetAttrs = captureImmutablePayload(
    {
      ...scopedAttrs(targetNode.attrs, scope.keys),
      [ORIGINAL_FORMATTING]: targetFormatting ?? null,
    },
    payloadContext,
  );
  if (targetAttrs.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: targetAttrs.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }
  const previousFormatting = captureImmutablePayload(baseFormatting, payloadContext);
  if (previousFormatting.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: previousFormatting.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }

  const acceptedNode = nodeWithAttrs(baseNode, targetAttrs.value);
  if (acceptedNode.isErr()) {
    return {
      status: "unsupported",
      issue: Object.freeze({
        reason: "non-reconstructable-property-change",
        reconstruction: "accepted",
        scope: scope.name,
        base,
        target,
      }),
    };
  }
  const acceptedState = capturedState(
    propertyStateValue(acceptedNode.value, scope),
    payloadContext,
  );
  if (acceptedState.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: acceptedState.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }
  if (acceptedState.value !== targetState.value) {
    return {
      status: "unsupported",
      issue: Object.freeze({
        reason: "non-reconstructable-property-change",
        reconstruction: "accepted",
        scope: scope.name,
        base,
        target,
      }),
    };
  }
  const rejectedNode = nodeWithAttrs(
    acceptedNode.value,
    scope.rejectPatch(baseFormatting, targetFormatting),
  );
  if (rejectedNode.isErr()) {
    return {
      status: "unsupported",
      issue: Object.freeze({
        reason: "non-reconstructable-property-change",
        reconstruction: "rejected",
        scope: scope.name,
        base,
        target,
      }),
    };
  }
  const rejectedState = capturedState(
    propertyStateValue(rejectedNode.value, scope),
    payloadContext,
  );
  if (rejectedState.isErr()) {
    return {
      status: "unsupported",
      issue: payloadFailureIssue({
        failure: rejectedState.error,
        scope: scope.name,
        base,
        target,
        maximum: payloadContext.maximum,
      }),
    };
  }
  if (rejectedState.value !== baseState.value) {
    return {
      status: "unsupported",
      issue: Object.freeze({
        reason: "non-reconstructable-property-change",
        reconstruction: "rejected",
        scope: scope.name,
        base,
        target,
      }),
    };
  }

  const instruction = scope.createInstruction(carrier, targetAttrs.value, previousFormatting.value);
  return { status: "ready", instruction };
};

const cellProjection = (cell: PMNode): string =>
  canonicalJson({
    structure: liveStructure(cell, "cell"),
    properties: propertyState(cell, CELL_SCOPE),
  });

const rowProjection = (row: PMNode): string => {
  const cells: string[] = [];
  row.forEach((cell) => {
    cells.push(cellProjection(cell));
  });
  return canonicalJson({ properties: propertyState(row, ROW_SCOPE), cells });
};

/** One line per table: complete modelled geometry, in document order. */
export const projectTableGeometry = (tables: readonly FolioStoryTable[]): string[] =>
  tables.map(({ node }) => {
    const rows: string[] = [];
    node.forEach((row) => {
      rows.push(rowProjection(row));
    });
    return canonicalJson({
      structure: liveStructure(node, "table"),
      properties: propertyState(node, TABLE_SCOPE),
      rows,
    });
  });

const PROGRAMS = new WeakSet<object>();
const LIMIT_NAMES = Object.freeze([
  "maxTables",
  "maxPairings",
  "maxVisitedNodes",
  "maxChanges",
  "maxPayloadUnits",
] as const satisfies readonly TableGeometryLimit[]);

const frozenCoordinate = ({
  tableIndex,
  rowIndex,
  cellIndex,
}: TableCellCoordinate): Readonly<TableCellCoordinate> =>
  Object.freeze({ tableIndex, rowIndex, cellIndex });

const validCoordinate = ({ tableIndex, rowIndex, cellIndex }: TableCellCoordinate): boolean =>
  Number.isSafeInteger(tableIndex) &&
  tableIndex >= 0 &&
  Number.isSafeInteger(rowIndex) &&
  rowIndex >= 0 &&
  Number.isSafeInteger(cellIndex) &&
  cellIndex >= 0;

const coordinateKey = ({ tableIndex, rowIndex, cellIndex }: TableCellCoordinate): string =>
  `${String(tableIndex)}:${String(rowIndex)}:${String(cellIndex)}`;

const rowKey = ({ tableIndex, rowIndex }: TableCellCoordinate): string =>
  `${String(tableIndex)}:${String(rowIndex)}`;

const compareCoordinates = (left: TableCellCoordinate, right: TableCellCoordinate): number =>
  left.tableIndex - right.tableIndex ||
  left.rowIndex - right.rowIndex ||
  left.cellIndex - right.cellIndex;

const unsupported = (issue: TableGeometryUnsupportedIssue): TableGeometryPreflightResult =>
  Object.freeze({ status: "unsupported", issue: Object.freeze(issue) });

const missing = ({
  reason,
  side,
  scope,
  coordinate,
}: {
  reason: "missing-table" | "missing-row" | "missing-cell" | "unexpected-node-role";
  side: TableGeometrySide;
  scope: TableGeometryScopeName;
  coordinate: Readonly<TableCellCoordinate>;
}): TableGeometryPreflightResult => unsupported({ reason, side, scope, coordinate });

type ChildPositionResult =
  | { readonly status: "ready"; readonly positions: readonly number[] }
  | { readonly status: "unsupported"; readonly issue: TableGeometryUnsupportedIssue };

const childPositions = (
  node: PMNode,
  start: number,
  limits: TableGeometryPreflightLimits,
  visited: { count: number },
): ChildPositionResult => {
  const actual = visited.count + node.childCount;
  if (actual > limits.maxVisitedNodes) {
    return {
      status: "unsupported",
      issue: {
        reason: "limit-exceeded",
        limit: "maxVisitedNodes",
        maximum: limits.maxVisitedNodes,
        actual,
      },
    };
  }
  visited.count = actual;
  const positions: number[] = [];
  let offset = start + 1;
  for (let index = 0; index < node.childCount; index++) {
    const child = node.child(index);
    positions.push(offset);
    offset += child.nodeSize;
  }
  return { status: "ready", positions: Object.freeze(positions) };
};

const nodeHasRole = (node: PMNode, role: string): boolean => node.type.spec["tableRole"] === role;

const nodeHasCellRole = (node: PMNode): boolean =>
  TABLE_CELL_NODE_ROLES.has(String(node.type.spec["tableRole"]));

type CanonicalPairing = {
  readonly base: Readonly<TableCellCoordinate>;
  readonly target: Readonly<TableCellCoordinate>;
};

type CapturedPairingsResult =
  | { readonly status: "ready"; readonly pairings: readonly CanonicalPairing[] }
  | { readonly status: "unsupported"; readonly issue: TableGeometryUnsupportedIssue };

const capturePairings = (pairings: readonly TableGeometryPairing[]): CapturedPairingsResult => {
  const baseCoordinates = new Set<string>();
  const targetCoordinates = new Set<string>();
  const baseTableTargets = new Map<number, number>();
  const targetTableBases = new Map<number, number>();
  const baseRowTargets = new Map<string, string>();
  const targetRowBases = new Map<string, string>();
  const captured: CanonicalPairing[] = [];

  for (const pairing of pairings) {
    const base = frozenCoordinate(pairing.base);
    const target = frozenCoordinate(pairing.target);
    if (!validCoordinate(base)) {
      return {
        status: "unsupported",
        issue: { reason: "invalid-coordinate", side: "base", coordinate: base },
      };
    }
    if (!validCoordinate(target)) {
      return {
        status: "unsupported",
        issue: { reason: "invalid-coordinate", side: "target", coordinate: target },
      };
    }
    const baseCoordinate = coordinateKey(base);
    const targetCoordinate = coordinateKey(target);
    if (baseCoordinates.has(baseCoordinate)) {
      return {
        status: "unsupported",
        issue: { reason: "duplicate-pairing", side: "base", coordinate: base },
      };
    }
    if (targetCoordinates.has(targetCoordinate)) {
      return {
        status: "unsupported",
        issue: { reason: "duplicate-pairing", side: "target", coordinate: target },
      };
    }
    baseCoordinates.add(baseCoordinate);
    targetCoordinates.add(targetCoordinate);

    const priorTargetTable = baseTableTargets.get(base.tableIndex);
    const priorBaseTable = targetTableBases.get(target.tableIndex);
    if (priorTargetTable !== undefined && priorTargetTable !== target.tableIndex) {
      return {
        status: "unsupported",
        issue: { reason: "conflicting-table-pairing", side: "base", coordinate: base },
      };
    }
    if (priorBaseTable !== undefined && priorBaseTable !== base.tableIndex) {
      return {
        status: "unsupported",
        issue: { reason: "conflicting-table-pairing", side: "target", coordinate: target },
      };
    }
    baseTableTargets.set(base.tableIndex, target.tableIndex);
    targetTableBases.set(target.tableIndex, base.tableIndex);

    const baseRow = rowKey(base);
    const targetRow = rowKey(target);
    const priorTargetRow = baseRowTargets.get(baseRow);
    const priorBaseRow = targetRowBases.get(targetRow);
    if (priorTargetRow !== undefined && priorTargetRow !== targetRow) {
      return {
        status: "unsupported",
        issue: { reason: "conflicting-row-pairing", side: "base", coordinate: base },
      };
    }
    if (priorBaseRow !== undefined && priorBaseRow !== baseRow) {
      return {
        status: "unsupported",
        issue: { reason: "conflicting-row-pairing", side: "target", coordinate: target },
      };
    }
    baseRowTargets.set(baseRow, targetRow);
    targetRowBases.set(targetRow, baseRow);
    captured.push(Object.freeze({ base, target }));
  }

  captured.sort(
    (left, right) =>
      compareCoordinates(left.target, right.target) || compareCoordinates(left.base, right.base),
  );
  return { status: "ready", pairings: Object.freeze(captured) };
};

type PreflightTableGeometryOptions = {
  readonly baseTables: readonly FolioStoryTable[];
  readonly targetTables: ReadonlyMap<number, PMNode>;
  readonly pairings: readonly TableGeometryPairing[];
  readonly limits?: TableGeometryPreflightLimits;
};

/**
 * Resolve every carrier and reversible property payload before touching a
 * transaction. The program owns its payload and never reads the target again.
 */
export const preflightTableGeometry = ({
  baseTables,
  targetTables,
  pairings,
  limits = DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS,
}: PreflightTableGeometryOptions): TableGeometryPreflightResult => {
  for (const name of LIMIT_NAMES) {
    if (!Number.isSafeInteger(limits[name]) || limits[name] < 0) {
      return unsupported({ reason: "invalid-limit", limit: name, actual: limits[name] });
    }
  }
  if (baseTables.length > limits.maxTables) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxTables",
      maximum: limits.maxTables,
      actual: baseTables.length,
    });
  }
  if (targetTables.size > limits.maxTables) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxTables",
      maximum: limits.maxTables,
      actual: targetTables.size,
    });
  }
  if (pairings.length > limits.maxPairings) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxPairings",
      maximum: limits.maxPairings,
      actual: pairings.length,
    });
  }

  const baseByIndex = new Map<number, FolioStoryTable>();
  const basePositions = new Set<number>();
  for (const table of baseTables) {
    if (!Number.isSafeInteger(table.index) || table.index < 0) {
      return unsupported({ reason: "invalid-table-index", side: "base", tableIndex: table.index });
    }
    if (baseByIndex.has(table.index)) {
      return unsupported({
        reason: "duplicate-table-index",
        side: "base",
        tableIndex: table.index,
      });
    }
    if (!Number.isSafeInteger(table.start) || table.start < 0) {
      return unsupported({ reason: "invalid-table-position", position: table.start });
    }
    if (basePositions.has(table.start)) {
      return unsupported({ reason: "duplicate-table-position", position: table.start });
    }
    if (!nodeHasRole(table.node, TABLE_NODE_ROLE)) {
      const coordinate = Object.freeze({ tableIndex: table.index, rowIndex: 0, cellIndex: 0 });
      return missing({
        reason: "unexpected-node-role",
        side: "base",
        scope: "table",
        coordinate,
      });
    }
    baseByIndex.set(table.index, table);
    basePositions.add(table.start);
  }

  const targetByIndex = new Map<number, PMNode>();
  for (const [tableIndex, node] of targetTables) {
    if (!Number.isSafeInteger(tableIndex) || tableIndex < 0) {
      return unsupported({ reason: "invalid-table-index", side: "target", tableIndex });
    }
    if (targetByIndex.has(tableIndex)) {
      return unsupported({ reason: "duplicate-table-index", side: "target", tableIndex });
    }
    if (!nodeHasRole(node, TABLE_NODE_ROLE)) {
      const coordinate = Object.freeze({ tableIndex, rowIndex: 0, cellIndex: 0 });
      return missing({
        reason: "unexpected-node-role",
        side: "target",
        scope: "table",
        coordinate,
      });
    }
    targetByIndex.set(tableIndex, node);
  }

  const capturedPairings = capturePairings(pairings);
  if (capturedPairings.status === "unsupported") return unsupported(capturedPairings.issue);

  const carriers: TableGeometryCarrierAssertion[] = [];
  const instructions: TableGeometryInstruction[] = [];
  const claimedPositions = new Map<number, TableGeometryScopeName>();
  const rowPositionsByTable = new Map<number, readonly number[]>();
  const cellPositionsByRow = new Map<string, readonly number[]>();
  const visited = { count: 0 };
  const payloadContext: PayloadCaptureContext = {
    maximum: limits.maxPayloadUnits,
    used: 0,
    active: new WeakSet(),
    captured: new WeakMap(),
  };

  const consider = <Scope extends TableGeometryScopeName>(
    baseNode: PMNode,
    targetNode: PMNode,
    position: number,
    scope: PropertyScope<Scope>,
    pairing: CanonicalPairing,
  ): TableGeometryPreflightResult | null => {
    const claimedScope = claimedPositions.get(position);
    if (claimedScope !== undefined) {
      return claimedScope === scope.name
        ? null
        : unsupported({ reason: "duplicate-table-position", position });
    }
    claimedPositions.set(position, scope.name);
    if (scope.name === "table") {
      const baseStructure = capturedState(liveStructure(baseNode, scope.name), payloadContext);
      if (baseStructure.isErr()) {
        return unsupported(
          payloadFailureIssue({
            failure: baseStructure.error,
            scope: scope.name,
            base: pairing.base,
            target: pairing.target,
            maximum: payloadContext.maximum,
          }),
        );
      }
      const targetStructure = capturedState(liveStructure(targetNode, scope.name), payloadContext);
      if (targetStructure.isErr()) {
        return unsupported(
          payloadFailureIssue({
            failure: targetStructure.error,
            scope: scope.name,
            base: pairing.base,
            target: pairing.target,
            maximum: payloadContext.maximum,
          }),
        );
      }
      if (baseStructure.value !== targetStructure.value) {
        return unsupported({
          reason: "non-reconstructable-structure-change",
          scope: "table",
          property: "column-widths",
          base: pairing.base,
          target: pairing.target,
        });
      }
    }
    const expectedLiveState = capturedLiveState(baseNode, scope, payloadContext);
    if (expectedLiveState.isErr()) {
      return unsupported(
        payloadFailureIssue({
          failure: expectedLiveState.error,
          scope: scope.name,
          base: pairing.base,
          target: pairing.target,
          maximum: payloadContext.maximum,
        }),
      );
    }
    const carrier = Object.freeze({
      scope: scope.name,
      position,
      expectedLiveState: expectedLiveState.value,
      base: pairing.base,
      target: pairing.target,
    });
    carriers.push(carrier);
    const result = propertyChangeFor({
      baseNode,
      targetNode,
      scope,
      carrier,
      payloadContext,
    });
    if (result.status === "unsupported") return unsupported(result.issue);
    if (result.status === "ready") {
      instructions.push(result.instruction);
      if (instructions.length > limits.maxChanges) {
        return unsupported({
          reason: "limit-exceeded",
          limit: "maxChanges",
          maximum: limits.maxChanges,
          actual: instructions.length,
        });
      }
    }
    return null;
  };

  for (const pairing of capturedPairings.pairings) {
    const baseTable = baseByIndex.get(pairing.base.tableIndex);
    if (!baseTable) {
      return missing({
        reason: "missing-table",
        side: "base",
        scope: "table",
        coordinate: pairing.base,
      });
    }
    const targetTable = targetByIndex.get(pairing.target.tableIndex);
    if (!targetTable) {
      return missing({
        reason: "missing-table",
        side: "target",
        scope: "table",
        coordinate: pairing.target,
      });
    }
    const baseRow = baseTable.node.maybeChild(pairing.base.rowIndex);
    if (!baseRow) {
      return missing({
        reason: "missing-row",
        side: "base",
        scope: "row",
        coordinate: pairing.base,
      });
    }
    const targetRow = targetTable.maybeChild(pairing.target.rowIndex);
    if (!targetRow) {
      return missing({
        reason: "missing-row",
        side: "target",
        scope: "row",
        coordinate: pairing.target,
      });
    }
    if (!nodeHasRole(baseRow, TABLE_ROW_NODE_ROLE)) {
      return missing({
        reason: "unexpected-node-role",
        side: "base",
        scope: "row",
        coordinate: pairing.base,
      });
    }
    if (!nodeHasRole(targetRow, TABLE_ROW_NODE_ROLE)) {
      return missing({
        reason: "unexpected-node-role",
        side: "target",
        scope: "row",
        coordinate: pairing.target,
      });
    }
    const baseCell = baseRow.maybeChild(pairing.base.cellIndex);
    if (!baseCell) {
      return missing({
        reason: "missing-cell",
        side: "base",
        scope: "cell",
        coordinate: pairing.base,
      });
    }
    const targetCell = targetRow.maybeChild(pairing.target.cellIndex);
    if (!targetCell) {
      return missing({
        reason: "missing-cell",
        side: "target",
        scope: "cell",
        coordinate: pairing.target,
      });
    }
    if (!nodeHasCellRole(baseCell)) {
      return missing({
        reason: "unexpected-node-role",
        side: "base",
        scope: "cell",
        coordinate: pairing.base,
      });
    }
    if (!nodeHasCellRole(targetCell)) {
      return missing({
        reason: "unexpected-node-role",
        side: "target",
        scope: "cell",
        coordinate: pairing.target,
      });
    }
    if (baseCell.type.name !== targetCell.type.name) {
      return unsupported({
        reason: "structural-cell-mismatch",
        property: "node-type",
        base: pairing.base,
        target: pairing.target,
      });
    }
    for (const property of ["colspan", "rowspan"] as const) {
      if ((baseCell.attrs[property] ?? 1) !== (targetCell.attrs[property] ?? 1)) {
        return unsupported({
          reason: "structural-cell-mismatch",
          property,
          base: pairing.base,
          target: pairing.target,
        });
      }
    }

    let rowPositions = rowPositionsByTable.get(baseTable.index);
    if (!rowPositions) {
      const positioned = childPositions(baseTable.node, baseTable.start, limits, visited);
      if (positioned.status === "unsupported") return unsupported(positioned.issue);
      rowPositions = positioned.positions;
      rowPositionsByTable.set(baseTable.index, rowPositions);
    }
    const rowPosition = rowPositions[pairing.base.rowIndex];
    if (rowPosition === undefined) {
      return missing({
        reason: "missing-row",
        side: "base",
        scope: "row",
        coordinate: pairing.base,
      });
    }
    const baseRowKey = rowKey(pairing.base);
    let cellPositions = cellPositionsByRow.get(baseRowKey);
    if (!cellPositions) {
      const positioned = childPositions(baseRow, rowPosition, limits, visited);
      if (positioned.status === "unsupported") return unsupported(positioned.issue);
      cellPositions = positioned.positions;
      cellPositionsByRow.set(baseRowKey, cellPositions);
    }
    const cellPosition = cellPositions[pairing.base.cellIndex];
    if (cellPosition === undefined) {
      return missing({
        reason: "missing-cell",
        side: "base",
        scope: "cell",
        coordinate: pairing.base,
      });
    }

    const tableResult = consider(
      baseTable.node,
      targetTable,
      baseTable.start,
      TABLE_SCOPE,
      pairing,
    );
    if (tableResult) return tableResult;
    const rowResult = consider(baseRow, targetRow, rowPosition, ROW_SCOPE, pairing);
    if (rowResult) return rowResult;
    const cellResult = consider(baseCell, targetCell, cellPosition, CELL_SCOPE, pairing);
    if (cellResult) return cellResult;
  }

  const scopeOrder = { table: 0, row: 1, cell: 2 } as const;
  const compareCarriers = (
    left: TableGeometryCarrierAssertion,
    right: TableGeometryCarrierAssertion,
  ): number =>
    compareCoordinates(left.target, right.target) ||
    scopeOrder[left.scope] - scopeOrder[right.scope] ||
    compareCoordinates(left.base, right.base);
  carriers.sort(compareCarriers);
  instructions.sort((left, right) => compareCarriers(left.carrier, right.carrier));
  const frozenCarriers = Object.freeze(carriers);
  const frozenInstructions = Object.freeze(instructions);
  const program = Object.freeze({
    [TABLE_GEOMETRY_PROGRAM_BRAND]: true as const,
    type: frozenInstructions.length === 0 ? ("unchanged" as const) : ("changes" as const),
    carriers: frozenCarriers,
    instructions: frozenInstructions,
    maxLivePayloadUnits: limits.maxPayloadUnits,
  });
  PROGRAMS.add(program);
  return Object.freeze({ status: "ready", program });
};

const capturedLiveStateForScope = (
  node: PMNode,
  scope: TableGeometryScopeName,
  context: PayloadCaptureContext,
): Result<string, PayloadCaptureFailure> => {
  switch (scope) {
    case "table":
      return capturedLiveState(node, TABLE_SCOPE, context);
    case "row":
      return capturedLiveState(node, ROW_SCOPE, context);
    case "cell":
      return capturedLiveState(node, CELL_SCOPE, context);
    default: {
      const unreachable: never = scope;
      return panic("Unhandled table geometry instruction", { scope: unreachable });
    }
  }
};

type ExecuteTableGeometryProgramOptions = {
  readonly tr: Transaction;
  readonly program: TableGeometryProgram;
  readonly revision: TableGeometryRevisionStamp;
};

/** Apply a preflighted program to one caller-owned transaction without dispatching it. */
export const executeTableGeometryProgram = ({
  tr,
  program,
  revision,
}: ExecuteTableGeometryProgramOptions): TableGeometryExecutionResult => {
  if (!PROGRAMS.has(program)) {
    return panic("A table geometry program must come from preflightTableGeometry");
  }
  if (
    typeof revision.author !== "string" ||
    typeof revision.date !== "string" ||
    revision.author.length + revision.date.length > 65_536 ||
    !Number.isSafeInteger(revision.idSeed) ||
    revision.idSeed < 0 ||
    revision.idSeed + program.instructions.length > Number.MAX_SAFE_INTEGER
  ) {
    return Object.freeze({
      status: "unsupported",
      issue: Object.freeze({ reason: "invalid-revision-stamp" }),
    });
  }

  const liveNodeAt = (position: number): PMNode | null =>
    position >= 0 && position < tr.doc.content.size ? tr.doc.nodeAt(position) : null;
  const payloadContext: PayloadCaptureContext = {
    maximum: program.maxLivePayloadUnits,
    used: 0,
    active: new WeakSet(),
    captured: new WeakMap(),
  };

  // Validate every carrier before the first step: stale input cannot apply a
  // prefix of the intended geometry program.
  for (const carrier of program.carriers) {
    const node = liveNodeAt(carrier.position);
    if (!node) {
      return Object.freeze({
        status: "unsupported",
        issue: Object.freeze({
          reason: "missing-live-node",
          scope: carrier.scope,
          position: carrier.position,
        }),
      });
    }
    const liveState = capturedLiveStateForScope(node, carrier.scope, payloadContext);
    if (liveState.isErr()) {
      return Object.freeze({
        status: "unsupported",
        issue: Object.freeze(
          liveState.error.type === "limit"
            ? {
                reason: "live-payload-limit" as const,
                maximum: payloadContext.maximum,
                actual: liveState.error.actual,
              }
            : {
                reason: "invalid-live-node" as const,
                scope: carrier.scope,
                position: carrier.position,
              },
        ),
      });
    }
    if (liveState.value !== carrier.expectedLiveState) {
      return Object.freeze({
        status: "unsupported",
        issue: Object.freeze({
          reason: "stale-live-node",
          scope: carrier.scope,
          position: carrier.position,
        }),
      });
    }
  }

  const receipts: TableGeometryRevisionReceipt[] = [];
  let revisionId = revision.idSeed;
  for (const instruction of program.instructions) {
    const { carrier } = instruction;
    const node = liveNodeAt(carrier.position);
    if (!node) return panic("A validated table geometry carrier disappeared before execution");
    const info = Object.freeze({ id: revisionId, author: revision.author, date: revision.date });
    const propertyChange = Object.freeze({
      type: instruction.change.changeType,
      info,
      ...(instruction.change.previousFormatting !== undefined && {
        previousFormatting: instruction.change.previousFormatting,
      }),
    });
    tr.setNodeMarkup(carrier.position, undefined, {
      ...node.attrs,
      ...instruction.targetAttrs,
      [instruction.change.changeAttr]: Object.freeze([propertyChange]),
    });
    receipts.push(
      Object.freeze({
        scope: carrier.scope,
        position: carrier.position,
        revisionId,
        base: carrier.base,
        target: carrier.target,
      }),
    );
    revisionId += 1;
  }
  if (receipts.length > 0) {
    // A property change writes no paragraph, so selective save needs one
    // structural signal to rewrite the owning story part.
    markStructuralChange(tr);
  }
  return Object.freeze({
    status: "executed",
    receipt: Object.freeze({
      type: "table-geometry-execution",
      startingRevisionId: revision.idSeed,
      nextRevisionId: revisionId,
      revisions: Object.freeze(receipts),
    }),
  });
};
