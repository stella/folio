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
export const FOLIO_DOCX_CAPABILITY_IDS = Object.freeze(["opaqueDrawing"] as const);

export type FolioDocxCapabilityHost = (typeof FOLIO_DOCX_CAPABILITY_HOSTS)[number];
export type FolioDocxCompatibilityHost = (typeof FOLIO_DOCX_COMPATIBILITY_HOSTS)[number];
export type FolioDocxProfile = (typeof FOLIO_DOCX_PROFILES)[number];
export type FolioDocxCompatibilityProfile = (typeof FOLIO_DOCX_COMPATIBILITY_PROFILES)[number];
export type FolioDocxCapabilityOperation = (typeof FOLIO_DOCX_CAPABILITY_OPERATIONS)[number];
export type FolioDocxSupportState = (typeof FOLIO_DOCX_SUPPORT_STATES)[number];
export type FolioDocxCapabilityId = (typeof FOLIO_DOCX_CAPABILITY_IDS)[number];

export type FolioDocxFeatureCapability = {
  readonly id: FolioDocxCapabilityId;
  readonly feature: "drawings";
  readonly hosts: readonly FolioDocxCapabilityHost[];
  readonly profiles: readonly FolioDocxProfile[];
  readonly support: Readonly<Record<FolioDocxCapabilityOperation, FolioDocxSupportState>>;
  readonly evidence: readonly {
    readonly type: "test";
    readonly path: string;
  }[];
};

const OPAQUE_DRAWING_CAPABILITY = Object.freeze({
  id: "opaqueDrawing",
  feature: "drawings",
  hosts: FOLIO_DOCX_CAPABILITY_HOSTS,
  profiles: Object.freeze(["transitional"]),
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

export const FOLIO_DOCX_CAPABILITY_MANIFEST = Object.freeze({
  version: FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION,
  capabilities: Object.freeze({
    opaqueDrawing: OPAQUE_DRAWING_CAPABILITY,
  }),
});

export class InvalidFolioDocxCapabilityIdError extends TaggedError(
  "InvalidFolioDocxCapabilityIdError",
)<{
  message: string;
  receivedId: unknown;
}>() {}

export const isFolioDocxCapabilityId = (value: unknown): value is FolioDocxCapabilityId =>
  value === "opaqueDrawing";

export const getFolioDocxCapability = (id: unknown): FolioDocxFeatureCapability => {
  if (!isFolioDocxCapabilityId(id)) {
    throw new InvalidFolioDocxCapabilityIdError({
      message: "Invalid DOCX capability id.",
      receivedId: id,
    });
  }
  return FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities[id];
};
import { TaggedError } from "better-result";
