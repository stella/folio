/** How a block identifier participates in comparison across revisions. */
export const FOLIO_CONTENT_IDENTITY_SEMANTICS = Object.freeze([
  "authoritative",
  "persistent-hint",
  "positional",
] as const);

export type FolioContentIdentitySemantics =
  (typeof FOLIO_CONTENT_IDENTITY_SEMANTICS)[number];

/** Identity and its semantics travel as one discriminated value. */
export type FolioContentIdentity = {
  readonly [Type in FolioContentIdentitySemantics]: {
    readonly type: Type;
    readonly id: string;
  };
}[FolioContentIdentitySemantics];

/** Total ownership map for the atomic identity record. @internal */
export const FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type" }),
  id: Object.freeze({ field: "id" }),
} as const satisfies SelfDescribingFieldMap<
  Extract<FolioContentIdentity, { type: "authoritative" }>,
  Record<never, never>
>);

/** Ergonomic caller input for a bounded, representation-neutral property value. */
export type FolioContentPropertyInputValue =
  | boolean
  | number
  | string
  | null
  | {
      readonly type: "array";
      readonly items: readonly FolioContentPropertyInputValue[];
    }
  | {
      readonly type: "object";
      readonly entries: FolioContentPropertyInput;
    };

/** Caller property entries; order is ignored and duplicate keys are invalid. */
export type FolioContentPropertyInput = readonly {
  readonly key: string;
  readonly value: FolioContentPropertyInputValue;
}[];

export const FOLIO_CONTENT_PROPERTY_ENTRY_FIELD_DESCRIPTORS = Object.freeze({
  key: Object.freeze({ field: "key" }),
  value: Object.freeze({ field: "value" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentPropertyInput[number],
  Record<never, never>
>);

export const FOLIO_CONTENT_PROPERTY_ARRAY_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type" }),
  items: Object.freeze({ field: "items" }),
} as const satisfies SelfDescribingFieldMap<
  Extract<FolioContentPropertyInputValue, { readonly type: "array" }>,
  Record<never, never>
>);

export const FOLIO_CONTENT_PROPERTY_OBJECT_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type" }),
  entries: Object.freeze({ field: "entries" }),
} as const satisfies SelfDescribingFieldMap<
  Extract<FolioContentPropertyInputValue, { readonly type: "object" }>,
  Record<never, never>
>);

/** Canonical owned property value used by comparison results. */
export type FolioContentPropertyValue =
  | boolean
  | number
  | string
  | null
  | {
      readonly type: "array";
      readonly items: readonly FolioContentPropertyValue[];
    }
  | {
      readonly type: "object";
      readonly entries: FolioContentPropertySet;
    };

/** One canonical object entry; sets are sorted by UTF-16 key order. */
export type FolioContentProperty = {
  readonly key: string;
  readonly value: FolioContentPropertyValue;
};

export type FolioContentPropertySet = readonly FolioContentProperty[];

export type FolioContentPropertyPresence =
  | { readonly type: "absent" }
  | { readonly type: "present"; readonly value: FolioContentPropertyValue };

/** Exact two-sided delta for one canonical property key. */
export type FolioContentPropertyChange = {
  readonly key: string;
  readonly base: FolioContentPropertyPresence;
  readonly revised: FolioContentPropertyPresence;
};

/** One enclosing structural container, ordered outermost to innermost in a block path. */
export type FolioContentContainerPathEntry = {
  readonly kind: string;
  readonly identity: FolioContentIdentity;
};

/** A supported zero-width inline structure at a UTF-16 text boundary. */
export type FolioContentStructuralBoundary = {
  readonly type: "pageBreak";
  readonly offset: number;
  readonly clear?: "all" | "left" | "right" | "none";
};

export type FolioContentInputRun = {
  readonly text: string;
  /** Fully resolved presentation for this range. */
  readonly effectiveFormatting?: FolioContentPropertyInput;
  /** Authored inline properties only; inherited values stay out of this set. */
  readonly authoredFormatting?: FolioContentPropertyInput;
};

/** Caller paragraph presentation with authored provenance kept separate from inheritance. */
export type FolioContentInputParagraphFormatting = {
  /** Fully resolved paragraph presentation, for renderers only. */
  readonly effective?: FolioContentPropertyInput;
  /** Authored paragraph properties only; inherited values stay out of this set. */
  readonly authored?: FolioContentPropertyInput;
};

/** Captured paragraph presentation whose two property sets are canonical and immutable. */
export type FolioContentParagraphFormatting = {
  readonly effective: FolioContentPropertySet;
  readonly authored: FolioContentPropertySet;
};

/** Captured run whose property sets have one canonical ordering and value algebra. */
export type FolioContentRun = {
  readonly text: string;
  readonly effectiveFormatting: FolioContentPropertySet;
  readonly authoredFormatting: FolioContentPropertySet;
};

/** Authored and effective presentation changes over one text-aligned range. */
export type FolioContentInlineFormattingChange = {
  /** Only these deltas may be lowered into authored transport properties. */
  readonly authored: readonly FolioContentPropertyChange[];
  /** Presentation-only deltas for renderers; these never imply authorship. */
  readonly effective: readonly FolioContentPropertyChange[];
};

/**
 * Where a block sits inside its innermost enclosing table. Every index is
 * zero-based and a merged physical cell occupies one `cellIndex` slot.
 */
export type FolioContentTableLocation = {
  readonly outerTableIdentity: FolioContentIdentity;
  readonly tableIdentity: FolioContentIdentity;
  readonly rowIdentity: FolioContentIdentity;
  readonly cellIdentity: FolioContentIdentity;
  readonly outerTableIndex: number;
  readonly tableIndex: number;
  readonly rowIndex: number;
  readonly cellIndex: number;
  readonly gridColumnIndex: number;
  readonly columnSpan: number;
  readonly rowSpan: number;
  readonly paragraphIndex: number;
};

/** Ergonomic representation-neutral input block for one ordered story. */
export type FolioContentInputBlock<Kind extends string = string> = {
  readonly identity: FolioContentIdentity;
  readonly kind: Kind;
  readonly text: string;
  readonly blockProperties?: FolioContentPropertyInput;
  readonly paragraphFormatting?: FolioContentInputParagraphFormatting;
  readonly runs?: readonly FolioContentInputRun[];
  readonly structuralBoundaries?: readonly FolioContentStructuralBoundary[];
  readonly table?: FolioContentTableLocation;
  readonly containerPath?: readonly FolioContentContainerPathEntry[];
};

/** Owned immutable block carried by the canonical comparison result. */
export type FolioContentBlock<Kind extends string = string> = {
  readonly identity: FolioContentIdentity;
  readonly kind: Kind;
  readonly text: string;
  readonly blockProperties: FolioContentPropertySet;
  readonly paragraphFormatting: FolioContentParagraphFormatting;
  readonly runs: readonly FolioContentRun[];
  readonly structuralBoundaries: readonly FolioContentStructuralBoundary[];
  readonly table?: FolioContentTableLocation;
  readonly containerPath: readonly FolioContentContainerPathEntry[];
};

/** Every caller-supplied block of one story, in document order. */
export type FolioContentSnapshot = {
  readonly blocks: readonly FolioContentInputBlock[];
};

export const FOLIO_CONTENT_SNAPSHOT_FIELD_DESCRIPTORS = Object.freeze({
  blocks: Object.freeze({ field: "blocks" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentSnapshot,
  Record<never, never>
>);

type SelfDescribingFieldMap<Value, Descriptor> = {
  [Field in keyof Value]-?: Descriptor & { readonly field: Field };
};

type FolioContentInputBlockFieldDescriptor =
  | {
      readonly role: "identity";
      readonly capture: "identity";
      readonly comparison: "none";
      readonly validation: "identity";
    }
  | {
      readonly role: "kind" | "text";
      readonly capture: "required-scalar";
      readonly comparison: "kind" | "none";
      readonly validation: "nonempty-string" | "text";
    }
  | {
      readonly role: "block-property" | "paragraph-format";
      readonly capture: "properties" | "paragraph-formatting";
      readonly comparison: "properties" | "paragraph";
      readonly validation: "properties";
    }
  | {
      readonly role: "inline-format";
      readonly capture: "runs";
      readonly comparison: "inline";
      readonly validation: "preview-runs";
    }
  | {
      readonly role: "structure";
      readonly capture: "boundaries" | "container" | "table";
      readonly comparison: "structural-boundaries" | "container" | "table";
      readonly validation: "structural-boundaries" | "container-path" | "table";
    };

/** Total ownership map for the public neutral input block. @internal */
export const FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS = Object.freeze({
  identity: Object.freeze({
    field: "identity",
    role: "identity",
    capture: "identity",
    comparison: "none",
    validation: "identity",
  }),
  kind: Object.freeze({
    field: "kind",
    role: "kind",
    capture: "required-scalar",
    comparison: "kind",
    validation: "nonempty-string",
  }),
  text: Object.freeze({
    field: "text",
    role: "text",
    capture: "required-scalar",
    comparison: "none",
    validation: "text",
  }),
  blockProperties: Object.freeze({
    field: "blockProperties",
    role: "block-property",
    capture: "properties",
    comparison: "properties",
    validation: "properties",
  }),
  paragraphFormatting: Object.freeze({
    field: "paragraphFormatting",
    role: "paragraph-format",
    capture: "paragraph-formatting",
    comparison: "paragraph",
    validation: "properties",
  }),
  runs: Object.freeze({
    field: "runs",
    role: "inline-format",
    capture: "runs",
    comparison: "inline",
    validation: "preview-runs",
  }),
  structuralBoundaries: Object.freeze({
    field: "structuralBoundaries",
    role: "structure",
    capture: "boundaries",
    comparison: "structural-boundaries",
    validation: "structural-boundaries",
  }),
  table: Object.freeze({
    field: "table",
    role: "structure",
    capture: "table",
    comparison: "table",
    validation: "table",
  }),
  containerPath: Object.freeze({
    field: "containerPath",
    role: "structure",
    capture: "container",
    comparison: "container",
    validation: "container-path",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentInputBlock,
  FolioContentInputBlockFieldDescriptor
>);

/** Total ownership map for authored and effective paragraph property sets. @internal */
export const FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS = Object.freeze({
  effective: Object.freeze({
    field: "effective",
    role: "effective-format",
    capture: "properties",
    validation: "properties",
  }),
  authored: Object.freeze({
    field: "authored",
    role: "authored-format",
    capture: "properties",
    validation: "properties",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentInputParagraphFormatting,
  {
    readonly role: "effective-format" | "authored-format";
    readonly capture: "properties";
    readonly validation: "properties";
  }
>);

/** Total ownership map for every caller-supplied run field. @internal */
export const FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS = Object.freeze({
  text: Object.freeze({
    field: "text",
    role: "text",
    capture: "required-scalar",
    validation: "text",
  }),
  effectiveFormatting: Object.freeze({
    field: "effectiveFormatting",
    role: "effective-format",
    capture: "properties",
    validation: "properties",
  }),
  authoredFormatting: Object.freeze({
    field: "authoredFormatting",
    role: "authored-format",
    capture: "properties",
    validation: "properties",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentInputRun,
  {
    readonly role: "text" | "effective-format" | "authored-format";
    readonly capture: "required-scalar" | "properties";
    readonly validation: "text" | "properties";
  }
>);

/** Total ownership map for table coordinates. @internal */
export const FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS = Object.freeze({
  outerTableIdentity: Object.freeze({ field: "outerTableIdentity", validation: "identity" }),
  tableIdentity: Object.freeze({ field: "tableIdentity", validation: "identity" }),
  rowIdentity: Object.freeze({ field: "rowIdentity", validation: "identity" }),
  cellIdentity: Object.freeze({ field: "cellIdentity", validation: "identity" }),
  outerTableIndex: Object.freeze({ field: "outerTableIndex", validation: "index" }),
  tableIndex: Object.freeze({ field: "tableIndex", validation: "index" }),
  rowIndex: Object.freeze({ field: "rowIndex", validation: "index" }),
  cellIndex: Object.freeze({ field: "cellIndex", validation: "index" }),
  gridColumnIndex: Object.freeze({ field: "gridColumnIndex", validation: "index" }),
  columnSpan: Object.freeze({ field: "columnSpan", validation: "span" }),
  rowSpan: Object.freeze({ field: "rowSpan", validation: "span" }),
  paragraphIndex: Object.freeze({ field: "paragraphIndex", validation: "index" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentTableLocation,
  { readonly validation: "identity" | "index" | "span" }
>);

export const FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS = Object.freeze({
  kind: Object.freeze({ field: "kind", validation: "nonempty-string" }),
  identity: Object.freeze({ field: "identity", validation: "identity" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentContainerPathEntry,
  { readonly validation: "identity" | "nonempty-string" }
>);

export const FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", validation: "page-break" }),
  offset: Object.freeze({ field: "offset", validation: "offset" }),
  clear: Object.freeze({ field: "clear", validation: "break-clear" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentStructuralBoundary,
  { readonly validation: "page-break" | "offset" | "break-clear" }
>);
