import type {
  FolioContentBlock,
  FolioContentContainerPathEntry,
  FolioContentIdentity,
  FolioContentIdentitySemantics,
  FolioContentTableLocation,
} from "./content-types";

const EMPTY_PROPERTIES = Object.freeze([]);
const EMPTY_RUNS = Object.freeze([]);
const EMPTY_BOUNDARIES = Object.freeze([]);
const EMPTY_CONTAINER_PATH = Object.freeze([]);
const EMPTY_PARAGRAPH_FORMATTING = Object.freeze({
  authored: EMPTY_PROPERTIES,
  effective: EMPTY_PROPERTIES,
});

export const contentIdentity = (
  id: string,
  type: FolioContentIdentitySemantics = "authoritative",
): FolioContentIdentity => Object.freeze({ type, id });

export type ContentBlockFixtureOptions = {
  readonly identityType?: FolioContentIdentitySemantics;
  readonly kind?: string;
  readonly table?: FolioContentTableLocation;
  readonly containerPath?: readonly FolioContentContainerPathEntry[];
};

/** Constructs the same complete immutable value shape produced by neutral capture. */
export const contentBlockFixture = (
  id: string,
  text: string,
  {
    identityType = "authoritative",
    kind = "paragraph",
    table,
    containerPath = EMPTY_CONTAINER_PATH,
  }: ContentBlockFixtureOptions = {},
): FolioContentBlock =>
  Object.freeze({
    identity: contentIdentity(id, identityType),
    kind,
    text,
    blockProperties: EMPTY_PROPERTIES,
    paragraphFormatting: EMPTY_PARAGRAPH_FORMATTING,
    runs: EMPTY_RUNS,
    structuralBoundaries: EMPTY_BOUNDARIES,
    ...(table === undefined ? {} : { table }),
    containerPath: Object.freeze([...containerPath]),
  });

type TableLocationFixtureOptions = Omit<
  FolioContentTableLocation,
  "outerTableIdentity" | "tableIdentity" | "rowIdentity" | "cellIdentity"
> & {
  readonly identityType?: FolioContentIdentitySemantics;
  readonly outerTableId?: string;
  readonly tableId?: string;
  readonly rowId?: string;
  readonly cellId?: string;
};

export const tableLocationFixture = ({
  identityType = "persistent-hint",
  outerTableId,
  tableId,
  rowId,
  cellId,
  ...coordinates
}: TableLocationFixtureOptions): FolioContentTableLocation =>
  Object.freeze({
    outerTableIdentity: contentIdentity(
      outerTableId ?? `outer-table:${String(coordinates.outerTableIndex)}`,
      identityType,
    ),
    tableIdentity: contentIdentity(
      tableId ?? `table:${String(coordinates.tableIndex)}`,
      identityType,
    ),
    rowIdentity: contentIdentity(rowId ?? `row:${String(coordinates.rowIndex)}`, identityType),
    cellIdentity: contentIdentity(
      cellId ?? `cell:${String(coordinates.rowIndex)}:${String(coordinates.cellIndex)}`,
      identityType,
    ),
    ...coordinates,
  });
