import { TaggedError } from "better-result";

export const FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION = 1 as const;

export const FOLIO_DOCX_CAPABILITY_HOSTS = Object.freeze(["browser", "server"] as const);
export const FOLIO_DOCX_COMPATIBILITY_HOSTS = Object.freeze([
  ...FOLIO_DOCX_CAPABILITY_HOSTS,
  "unknown",
] as const);
export const FOLIO_DOCX_PROFILES = Object.freeze(["strict", "transitional"] as const);
export const FOLIO_DOCX_COMPATIBILITY_PROFILES = Object.freeze([
  ...FOLIO_DOCX_PROFILES,
  "unknown",
] as const);
export const FOLIO_DOCX_CAPABILITY_OPERATIONS = Object.freeze([
  "create",
  "edit",
  "preserve",
  "read",
  "render",
] as const);
export const FOLIO_DOCX_SUPPORT_STATES = Object.freeze([
  "supported",
  "partial",
  "unsupported",
] as const);
export const FOLIO_DOCX_CAPABILITY_IDS = Object.freeze([
  "comments",
  "opaqueDrawing",
  "paragraphs",
  "tables",
  "trackedChanges",
] as const);

export type FolioDocxCapabilityHost = (typeof FOLIO_DOCX_CAPABILITY_HOSTS)[number];
export type FolioDocxCompatibilityHost = (typeof FOLIO_DOCX_COMPATIBILITY_HOSTS)[number];
export type FolioDocxProfile = (typeof FOLIO_DOCX_PROFILES)[number];
export type FolioDocxCompatibilityProfile = (typeof FOLIO_DOCX_COMPATIBILITY_PROFILES)[number];
export type FolioDocxCapabilityOperation = (typeof FOLIO_DOCX_CAPABILITY_OPERATIONS)[number];
export type FolioDocxSupportState = (typeof FOLIO_DOCX_SUPPORT_STATES)[number];
export type FolioDocxCapabilityId = (typeof FOLIO_DOCX_CAPABILITY_IDS)[number];

export type FolioDocxFeatureCapability = {
  readonly id: FolioDocxCapabilityId;
  readonly feature: "comments" | "drawings" | "paragraphs" | "revisions" | "tables";
  readonly hosts: readonly FolioDocxCapabilityHost[];
  readonly profiles: readonly FolioDocxProfile[];
  readonly support: Readonly<Record<FolioDocxCapabilityOperation, FolioDocxSupportState>>;
  readonly evidence: readonly {
    readonly type: "test";
    readonly path: string;
  }[];
};

export type FolioDocxCapabilityManifest = {
  readonly version: typeof FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION;
  readonly capabilities: Readonly<Record<FolioDocxCapabilityId, FolioDocxFeatureCapability>>;
};

const TRANSITIONAL_PROFILE_COVERAGE = Object.freeze(["transitional"] as const);
const STRUCTURED_FEATURE_SUPPORT = Object.freeze({
  create: "supported",
  edit: "supported",
  preserve: "supported",
  read: "supported",
  render: "partial",
} as const satisfies Readonly<Record<FolioDocxCapabilityOperation, FolioDocxSupportState>>);

const COMMENTS_CAPABILITY = Object.freeze({
  id: "comments",
  feature: "comments",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: TRANSITIONAL_PROFILE_COVERAGE,
  support: STRUCTURED_FEATURE_SUPPORT,
  evidence: Object.freeze([
    {
      type: "test",
      path: "packages/core/src/docx/commentReplyThreads.test.ts",
    },
    {
      type: "test",
      path: "packages/core/src/prosemirror/commands/comments.test.ts",
    },
  ]),
} as const satisfies FolioDocxFeatureCapability);

const OPAQUE_DRAWING_CAPABILITY = Object.freeze({
  id: "opaqueDrawing",
  feature: "drawings",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: TRANSITIONAL_PROFILE_COVERAGE,
  support: Object.freeze({
    create: "unsupported",
    edit: "unsupported",
    preserve: "supported",
    read: "supported",
    render: "partial",
  }),
  evidence: Object.freeze([
    {
      type: "test",
      path: "packages/core/src/docx/compatibility.test.ts",
    },
    {
      type: "test",
      path: "packages/core/src/docx/rezip.test.ts",
    },
  ]),
} as const satisfies FolioDocxFeatureCapability);

const PARAGRAPHS_CAPABILITY = Object.freeze({
  id: "paragraphs",
  feature: "paragraphs",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: TRANSITIONAL_PROFILE_COVERAGE,
  support: STRUCTURED_FEATURE_SUPPORT,
  evidence: Object.freeze([
    {
      type: "test",
      path: "packages/core/src/docx/paragraphParser.test.ts",
    },
    {
      type: "test",
      path: "packages/core/src/docx/serializer/paragraphSerializer.test.ts",
    },
  ]),
} as const satisfies FolioDocxFeatureCapability);

const TABLES_CAPABILITY = Object.freeze({
  id: "tables",
  feature: "tables",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: TRANSITIONAL_PROFILE_COVERAGE,
  support: STRUCTURED_FEATURE_SUPPORT,
  evidence: Object.freeze([
    {
      type: "test",
      path: "packages/core/src/docx/tableParser.test.ts",
    },
    {
      type: "test",
      path: "packages/core/src/prosemirror/conversion/tableBordersRoundtrip.test.ts",
    },
  ]),
} as const satisfies FolioDocxFeatureCapability);

const TRACKED_CHANGES_CAPABILITY = Object.freeze({
  id: "trackedChanges",
  feature: "revisions",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: TRANSITIONAL_PROFILE_COVERAGE,
  support: STRUCTURED_FEATURE_SUPPORT,
  evidence: Object.freeze([
    {
      type: "test",
      path: "packages/core/src/redline.test.ts",
    },
    {
      type: "test",
      path: "packages/core/src/prosemirror/plugins/suggestionMode.test.ts",
    },
  ]),
} as const satisfies FolioDocxFeatureCapability);

const CAPABILITY_MANIFEST = Object.freeze({
  version: FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION,
  capabilities: Object.freeze({
    comments: COMMENTS_CAPABILITY,
    opaqueDrawing: OPAQUE_DRAWING_CAPABILITY,
    paragraphs: PARAGRAPHS_CAPABILITY,
    tables: TABLES_CAPABILITY,
    trackedChanges: TRACKED_CHANGES_CAPABILITY,
  }),
} as const satisfies FolioDocxCapabilityManifest);

export const FOLIO_DOCX_CAPABILITY_MANIFEST: FolioDocxCapabilityManifest = CAPABILITY_MANIFEST;

export class InvalidFolioDocxCapabilityIdError extends TaggedError(
  "InvalidFolioDocxCapabilityIdError",
)<{
  message: string;
  receivedId: unknown;
}>() {}

export const isFolioDocxCapabilityId = (value: unknown): value is FolioDocxCapabilityId =>
  FOLIO_DOCX_CAPABILITY_IDS.some((id) => id === value);

export const getFolioDocxCapability = (id: unknown): FolioDocxFeatureCapability => {
  if (!isFolioDocxCapabilityId(id)) {
    throw new InvalidFolioDocxCapabilityIdError({
      message: "Invalid DOCX capability id.",
      receivedId: id,
    });
  }
  return FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities[id];
};
