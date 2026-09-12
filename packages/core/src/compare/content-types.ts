/** How a block identifier participates in comparison across revisions. */
export const FOLIO_CONTENT_IDENTITY_SEMANTICS = Object.freeze([
  "authoritative",
  "persistent-hint",
  "positional",
] as const);

export type FolioContentIdentitySemantics = (typeof FOLIO_CONTENT_IDENTITY_SEMANTICS)[number];

/** Identity and its semantics travel as one discriminated value. */
export type FolioContentIdentity = {
  readonly [Type in FolioContentIdentitySemantics]: {
    readonly type: Type;
    readonly id: string;
  };
}[FolioContentIdentitySemantics];

/** Total ownership map for the atomic identity record. @internal */
export const FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", verification: "exact" }),
  id: Object.freeze({ field: "id", verification: "exact" }),
} as const satisfies SelfDescribingFieldMap<
  Extract<FolioContentIdentity, { type: "authoritative" }>,
  { readonly verification: "exact" }
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

/** One paragraph container occurrence on one side of a comparison. */
export type FolioContentContainerOccurrence =
  | {
      readonly type: "body";
      readonly containerPath: readonly FolioContentContainerPathEntry[];
      readonly end: "paragraph" | "structuralSibling";
    }
  | {
      readonly type: "tableCell";
      readonly containerPath: readonly FolioContentContainerPathEntry[];
      readonly end: "paragraph";
      /** Cell identity and placement; paragraph order is deliberately excluded. */
      readonly table: Omit<FolioContentTableLocation, "paragraphIndex">;
    };

/**
 * One alignment-owned container correspondence. A one-sided occurrence is
 * retained explicitly instead of being guessed into a neighbouring container.
 */
export type FolioContentContainerAlignment =
  | {
      readonly type: "paired";
      readonly id: number;
      readonly base: FolioContentContainerOccurrence;
      readonly revised: FolioContentContainerOccurrence;
    }
  | {
      readonly type: "baseOnly";
      readonly id: number;
      readonly base: FolioContentContainerOccurrence;
      readonly revised: null;
    }
  | {
      readonly type: "revisedOnly";
      readonly id: number;
      readonly base: null;
      readonly revised: FolioContentContainerOccurrence;
    };

export type FolioContentPairedContainerAlignment = Extract<
  FolioContentContainerAlignment,
  { readonly type: "paired" }
>;

export type FolioContentBaseContainerAlignment = Extract<
  FolioContentContainerAlignment,
  { readonly type: "paired" | "baseOnly" }
>;

export type FolioContentRevisedContainerAlignment = Extract<
  FolioContentContainerAlignment,
  { readonly type: "paired" | "revisedOnly" }
>;

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

/**
 * A target-side paragraph boundary proved inside one alignment-owned
 * container. An unanchored container never borrows a paragraph elsewhere.
 */
export type FolioContentParagraphInsertionBoundary<
  Block extends FolioContentBlock = FolioContentBlock,
> =
  | {
      readonly type: "beforeParagraph" | "afterParagraph";
      readonly paragraph: Block;
      readonly containerAlignment: FolioContentPairedContainerAlignment;
    }
  | {
      readonly type: "unanchoredContainer";
      readonly containerAlignment: FolioContentRevisedContainerAlignment;
    };

/**
 * A source-side paragraph removal boundary proved inside one alignment-owned
 * container. Terminal removal owns the predecessor whose mark is removed and
 * the revised paragraph whose properties the retained carrier must acquire.
 */
export type FolioContentParagraphRemovalBoundary<
  Block extends FolioContentBlock = FolioContentBlock,
> =
  | {
      readonly type: "successorParagraph";
      readonly successor: Block;
      readonly containerAlignment: FolioContentBaseContainerAlignment;
    }
  | {
      readonly type: "terminalPredecessor";
      readonly predecessor: Block;
      readonly targetCarrier: Block;
      readonly containerAlignment: FolioContentPairedContainerAlignment;
    }
  | {
      readonly type: "unanchoredContainer";
      readonly containerAlignment: FolioContentBaseContainerAlignment;
    };

/** Every caller-supplied block of one story, in document order. */
export type FolioContentSnapshot = {
  readonly blocks: readonly FolioContentInputBlock[];
};

export const FOLIO_CONTENT_SNAPSHOT_FIELD_DESCRIPTORS = Object.freeze({
  blocks: Object.freeze({ field: "blocks" }),
} as const satisfies SelfDescribingFieldMap<FolioContentSnapshot, Record<never, never>>);

type SelfDescribingFieldMap<Value, Descriptor> = {
  [Field in keyof Value]-?: Descriptor & { readonly field: Field };
};

type FolioContentInputBlockFieldDescriptor = {
  readonly verification: "container" | "exact" | "nested" | "transport-identity";
} & (
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
    }
);

/** Total ownership map for the public neutral input block. @internal */
export const FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS = Object.freeze({
  identity: Object.freeze({
    field: "identity",
    role: "identity",
    capture: "identity",
    comparison: "none",
    validation: "identity",
    verification: "transport-identity",
  }),
  kind: Object.freeze({
    field: "kind",
    role: "kind",
    capture: "required-scalar",
    comparison: "kind",
    validation: "nonempty-string",
    verification: "exact",
  }),
  text: Object.freeze({
    field: "text",
    role: "text",
    capture: "required-scalar",
    comparison: "none",
    validation: "text",
    verification: "exact",
  }),
  blockProperties: Object.freeze({
    field: "blockProperties",
    role: "block-property",
    capture: "properties",
    comparison: "properties",
    validation: "properties",
    verification: "exact",
  }),
  paragraphFormatting: Object.freeze({
    field: "paragraphFormatting",
    role: "paragraph-format",
    capture: "paragraph-formatting",
    comparison: "paragraph",
    validation: "properties",
    verification: "nested",
  }),
  runs: Object.freeze({
    field: "runs",
    role: "inline-format",
    capture: "runs",
    comparison: "inline",
    validation: "preview-runs",
    verification: "nested",
  }),
  structuralBoundaries: Object.freeze({
    field: "structuralBoundaries",
    role: "structure",
    capture: "boundaries",
    comparison: "structural-boundaries",
    validation: "structural-boundaries",
    verification: "nested",
  }),
  table: Object.freeze({
    field: "table",
    role: "structure",
    capture: "table",
    comparison: "table",
    validation: "table",
    verification: "nested",
  }),
  containerPath: Object.freeze({
    field: "containerPath",
    role: "structure",
    capture: "container",
    comparison: "container",
    validation: "container-path",
    verification: "container",
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
    verification: "exact",
  }),
  authored: Object.freeze({
    field: "authored",
    role: "authored-format",
    capture: "properties",
    validation: "properties",
    verification: "exact",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentInputParagraphFormatting,
  {
    readonly role: "effective-format" | "authored-format";
    readonly capture: "properties";
    readonly validation: "properties";
    readonly verification: "exact";
  }
>);

/** Total ownership map for every caller-supplied run field. @internal */
export const FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS = Object.freeze({
  text: Object.freeze({
    field: "text",
    role: "text",
    capture: "required-scalar",
    validation: "text",
    verification: "exact",
  }),
  effectiveFormatting: Object.freeze({
    field: "effectiveFormatting",
    role: "effective-format",
    capture: "properties",
    validation: "properties",
    verification: "exact",
  }),
  authoredFormatting: Object.freeze({
    field: "authoredFormatting",
    role: "authored-format",
    capture: "properties",
    validation: "properties",
    verification: "exact",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentInputRun,
  {
    readonly role: "text" | "effective-format" | "authored-format";
    readonly capture: "required-scalar" | "properties";
    readonly validation: "text" | "properties";
    readonly verification: "exact";
  }
>);

/** Total ownership map for table coordinates. @internal */
export const FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS = Object.freeze({
  outerTableIdentity: Object.freeze({
    field: "outerTableIdentity",
    validation: "identity",
    verification: "transport-identity",
  }),
  tableIdentity: Object.freeze({
    field: "tableIdentity",
    validation: "identity",
    verification: "transport-identity",
  }),
  rowIdentity: Object.freeze({
    field: "rowIdentity",
    validation: "identity",
    verification: "transport-identity",
  }),
  cellIdentity: Object.freeze({
    field: "cellIdentity",
    validation: "identity",
    verification: "transport-identity",
  }),
  outerTableIndex: Object.freeze({
    field: "outerTableIndex",
    validation: "index",
    verification: "exact",
  }),
  tableIndex: Object.freeze({ field: "tableIndex", validation: "index", verification: "exact" }),
  rowIndex: Object.freeze({ field: "rowIndex", validation: "index", verification: "exact" }),
  cellIndex: Object.freeze({ field: "cellIndex", validation: "index", verification: "exact" }),
  gridColumnIndex: Object.freeze({
    field: "gridColumnIndex",
    validation: "index",
    verification: "exact",
  }),
  columnSpan: Object.freeze({ field: "columnSpan", validation: "span", verification: "exact" }),
  rowSpan: Object.freeze({ field: "rowSpan", validation: "span", verification: "exact" }),
  paragraphIndex: Object.freeze({
    field: "paragraphIndex",
    validation: "index",
    verification: "exact",
  }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentTableLocation,
  {
    readonly validation: "identity" | "index" | "span";
    readonly verification: "exact" | "transport-identity";
  }
>);

export const FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS = Object.freeze({
  kind: Object.freeze({ field: "kind", validation: "nonempty-string", verification: "exact" }),
  identity: Object.freeze({ field: "identity", validation: "identity", verification: "exact" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentContainerPathEntry,
  { readonly validation: "identity" | "nonempty-string"; readonly verification: "exact" }
>);

export const FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", validation: "page-break", verification: "exact" }),
  offset: Object.freeze({ field: "offset", validation: "offset", verification: "exact" }),
  clear: Object.freeze({ field: "clear", validation: "break-clear", verification: "exact" }),
} as const satisfies SelfDescribingFieldMap<
  FolioContentStructuralBoundary,
  {
    readonly validation: "page-break" | "offset" | "break-clear";
    readonly verification: "exact";
  }
>);
