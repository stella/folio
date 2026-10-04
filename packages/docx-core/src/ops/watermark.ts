/** Document watermark ownership and forward-only header coverage. */
import { Result } from "better-result";
import type {
  BlockContent,
  Document,
  DocumentBody,
  HeaderFooter,
  HeaderFooterType,
  Paragraph,
  SectionProperties,
  Watermark,
} from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { normalizeCanonicalWatermark } from "../model/watermarkDefaults";
import { withBodyContent } from "./blocks";
import { idKey, isParaId, packageParagraphIds } from "./ids";
import { leafSpans } from "./leaves";
import { isRemovedRevisionNode } from "./review";
import { cloneModel } from "./modelClone";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { storyLifecycleEdit } from "./storyLifecycle";
import { DOCUMENT_OP_TYPES, type SetDocumentWatermarkOp } from "./types";

const refuse = (message: string) =>
  Result.err(
    new DocumentOpRefusal({
      opType: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message,
    }),
  );
const isPayloadRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const watermarkPayloadIssue = (value: unknown): string | undefined => {
  if (!isPayloadRecord(value)) return "A watermark must be a record.";
  switch (value["kind"]) {
    case "text": {
      const text = value["text"];
      if (typeof text !== "string" || text.length === 0 || hasIllegalXmlCharacters(text))
        return "A text watermark needs nonempty XML text.";
      const font = value["font"];
      if (
        font !== undefined &&
        (typeof font !== "string" || hasIllegalXmlCharacters(font) || font.includes(";"))
      )
        return "A watermark font must be a valid XML and CSS font string.";
      const color = value["color"];
      if (
        color !== undefined &&
        (typeof color !== "string" || (color !== "auto" && !/^[0-9a-f]{6}$/iu.test(color)))
      )
        return "A watermark color must be six hex digits or auto.";
      const diagonal = value["diagonal"];
      if (diagonal !== undefined && typeof diagonal !== "boolean")
        return "Watermark diagonal must be boolean.";
      const opacity = value["opacity"];
      if (
        opacity !== undefined &&
        (typeof opacity !== "number" || !Number.isFinite(opacity) || opacity < 0 || opacity > 1)
      )
        return "Watermark opacity must be finite and between zero and one.";
      return undefined;
    }
    case "picture": {
      const imageRId = value["imageRId"];
      if (
        typeof imageRId !== "string" ||
        imageRId.trim() === "" ||
        hasIllegalXmlCharacters(imageRId)
      )
        return "A picture watermark needs an XML relationship identity.";
      const target = value["imageTarget"];
      if (typeof target !== "string" || target.trim() === "" || hasIllegalXmlCharacters(target))
        return "A watermark image target must be an XML string.";
      for (const key of ["imageTargetExternal", "washout"]) {
        const field = value[key];
        if (field !== undefined && typeof field !== "boolean")
          return "Watermark target mode and washout must be boolean.";
      }
      for (const key of ["widthPt", "heightPt", "scale"]) {
        const field = value[key];
        if (
          field !== undefined &&
          (typeof field !== "number" || !Number.isFinite(field) || field <= 0)
        )
          return "Watermark dimensions and scale must be finite and positive.";
      }
      return undefined;
    }
    default:
      return "The watermark kind is invalid.";
  }
};
const identityPayloadIssue = (
  entries: readonly unknown[],
  coverage: "coverage" | "hosts",
): string | undefined => {
  for (const entry of entries) {
    if (
      !isPayloadRecord(entry) ||
      typeof entry["rId"] !== "string" ||
      typeof entry["paraId"] !== "string"
    )
      return "Watermark identities must be records with string identities.";
    if (hasIllegalXmlCharacters(entry["rId"]))
      return "Watermark relationship identities must be XML strings.";
    if (
      coverage === "coverage" &&
      entry["type"] !== "default" &&
      entry["type"] !== "first" &&
      entry["type"] !== "even"
    )
      return "The watermark coverage variant is invalid.";
  }
  return undefined;
};
const sectionSlots = (document: Document) => {
  const slots = document.package.document.content.flatMap((block, index) =>
    block.type === "paragraph" && block.sectionProperties !== undefined
      ? [{ properties: block.sectionProperties, blockIndex: index }]
      : [],
  );
  return [
    ...slots,
    { properties: document.package.document.finalSectionProperties ?? {}, blockIndex: null },
  ];
};
const coveragePlan = (document: Document) => {
  const slots = sectionSlots(document);
  const needed: { type: HeaderFooterType; sectionIndex: number }[] = [];
  const types: HeaderFooterType[] = ["default", "first"];
  if (document.package.settings?.evenAndOddHeaders === true) types.push("even");
  for (const type of types) {
    let inherited = false;
    for (const [sectionIndex, slot] of slots.entries()) {
      if (
        slot.properties.headerReferences?.some(
          (reference) => reference.type === type && document.package.headers?.has(reference.rId),
        )
      )
        inherited = true;
      if (inherited || (type === "first" && slot.properties.titlePg !== true)) continue;
      needed.push({ type, sectionIndex });
      inherited = true;
    }
  }
  return { slots, needed };
};
/** Each absent variant is created at its first non-inheriting section, at most once. */
export const planDocumentWatermarkCoverage = (document: Document): HeaderFooterType[] =>
  coveragePlan(document).needed.map(({ type }) => type);

/** Legacy visible-text semantics, rather than physical annotation atom width. */
export const isEmptyWatermarkHostParagraph = (paragraph: Paragraph): boolean => {
  const visible = leafSpans(paragraph.content)
    .filter(({ ancestors }) => !ancestors.some(isRemovedRevisionNode))
    .map(({ node }) => {
      switch (node.type) {
        case "text":
          return node.text;
        case "tab":
          return "\t";
        case "break":
          return node.breakType === "page" ? "\f" : "\n";
        case "softHyphen":
          return "\u00ad";
        case "noBreakHyphen":
          return "\u2011";
        case "mathEquation":
          return node.plainText ?? "";
        case "preservedInline":
          return node.text;
        default:
          return "";
      }
    })
    .join("");
  return (
    visible.trim() === "" &&
    !paragraph.content.some(
      (node) =>
        node.type === "run" &&
        node.content.some((child) => child.type === "drawing" || child.type === "shape"),
    )
  );
};
const watermarkHost = (header: HeaderFooter): Paragraph | undefined => {
  const host =
    header.watermarkBlockIndex === undefined
      ? undefined
      : header.content.at(header.watermarkBlockIndex);
  return host?.type === "paragraph" &&
    host.paraId !== undefined &&
    isParaId(host.paraId) &&
    isEmptyWatermarkHostParagraph(host) &&
    leafSpans(host.content).every(
      ({ node, ancestors }) =>
        node.type === "text" && ancestors.every((ancestor) => ancestor.type === "run"),
    )
    ? host
    : undefined;
};
/** Existing header parts that need a fresh explicit canonical watermark host. */
export const planDocumentWatermarkHosts = (document: Document): string[] =>
  [...(document.package.headers ?? [])].flatMap(([rId, header]) =>
    watermarkHost(header) ? [] : [rId],
  );

export type WatermarkHeaderCoverage = {
  type: HeaderFooterType;
  rId: string;
  content: BlockContent[];
};
type WatermarkCoverageOptions = {
  document: Document;
  authority: "legacy" | "canonical";
  watermark: Watermark;
  coverage: readonly WatermarkHeaderCoverage[];
};
type CanonicalSectionHeadersOptions = {
  body: DocumentBody;
  headers: ReadonlyMap<string, HeaderFooter>;
  propertiesBySection?: ReadonlyMap<number, SectionProperties>;
};
const syncCanonicalSectionHeaders = ({
  body,
  headers,
  propertiesBySection,
}: CanonicalSectionHeadersOptions): DocumentBody => {
  if (!body.sections) return body;
  const inherited = new Map<HeaderFooterType, HeaderFooter>();
  return {
    ...body,
    sections: body.sections.map((section, index) => {
      const properties = propertiesBySection?.get(index) ?? section.properties;
      for (const reference of properties.headerReferences ?? []) {
        const header = headers.get(reference.rId);
        if (header) inherited.set(reference.type, header);
        else inherited.delete(reference.type);
      }
      return { ...section, properties, headers: new Map(inherited) };
    }),
  };
};
/** Coverage-only legacy boundary: existing watermark/raw source payloads remain intact. */
export const ensureDocumentWatermarkHeaderCoverage = ({
  document,
  authority,
  watermark,
  coverage,
}: WatermarkCoverageOptions): Result<Document, DocumentOpRefusal> => {
  const { slots, needed } = coveragePlan(document);
  if (
    coverage.length !== needed.length ||
    new Set(coverage.map(({ type }) => type)).size !== coverage.length ||
    needed.some(({ type }) => !coverage.some((entry) => entry.type === type))
  )
    return refuse("Watermark coverage must name exactly the missing header variants.");
  if (needed.length === 0) return Result.ok(document);
  const used = new Set([
    ...(document.package.relationships?.keys() ?? []),
    ...(document.package.headers?.keys() ?? []),
    ...(document.package.footers?.keys() ?? []),
  ]);
  for (const entry of coverage) {
    if (entry.rId.trim() === "" || used.has(entry.rId))
      return refuse("Watermark coverage needs fresh package relationship identities.");
    used.add(entry.rId);
  }
  const headers = new Map(document.package.headers);
  const changes = new Map<number, SectionProperties>();
  for (const { type, sectionIndex } of needed) {
    const entry = coverage.find((candidate) => candidate.type === type);
    const slot = slots.at(sectionIndex);
    if (!entry || !slot) return refuse("The watermark coverage plan is incomplete.");
    headers.set(entry.rId, {
      type: "header",
      hdrFtrType: type,
      content: cloneModel(entry.content),
      watermark: cloneModel(watermark),
      ...(authority === "canonical" ? { watermarkBlockIndex: 0 } : {}),
    });
    const properties = changes.get(sectionIndex) ?? slot.properties;
    changes.set(sectionIndex, {
      ...properties,
      headerReferences: [...(properties.headerReferences ?? []), { type, rId: entry.rId }],
    });
  }
  const body = document.package.document;
  const content = [...body.content];
  let nextBody = { ...body };
  for (const [sectionIndex, properties] of changes) {
    const slot = slots.at(sectionIndex);
    if (slot?.blockIndex === null) nextBody.finalSectionProperties = properties;
    else if (slot) {
      const block = content.at(slot.blockIndex);
      if (block?.type === "paragraph")
        content[slot.blockIndex] = { ...block, sectionProperties: properties };
    }
  }
  if ([...changes.keys()].some((sectionIndex) => slots.at(sectionIndex)?.blockIndex !== null))
    nextBody = withBodyContent(nextBody, content);
  if (authority === "canonical")
    nextBody = syncCanonicalSectionHeaders({
      body: nextBody,
      headers,
      propertiesBySection: changes,
    });
  return Result.ok({ ...document, package: { ...document.package, headers, document: nextBody } });
};
type SetWatermarkOptions = {
  document: Document;
  change: SetDocumentWatermarkOp["change"];
  coverage: readonly WatermarkHeaderCoverage[];
} & ({ authority: "legacy" } | { authority: "canonical"; hosts: SetDocumentWatermarkOp["hosts"] });
/** One model algorithm serves canonical ops and the legacy immutable wrapper. */
export const setDocumentWatermarkWithCoverage = (
  options: SetWatermarkOptions,
): Result<Document, DocumentOpRefusal> => {
  const { document, change, coverage, authority } = options;
  if (change.kind === "remove" && coverage.length !== 0)
    return refuse("Removing a watermark cannot create coverage.");
  const headers = document.package.headers;
  const nextHeaders = new Map<string, HeaderFooter>();
  for (const [rId, header] of headers ?? []) {
    const next = { ...header };
    const hostIndex = next.watermarkBlockIndex;
    const rawHost =
      next.rawWatermarkXml !== undefined && hostIndex !== undefined
        ? next.content.at(hostIndex)
        : undefined;
    const host = authority === "canonical" ? watermarkHost(header) : rawHost;
    if (
      (authority === "legacy" || change.kind === "remove") &&
      hostIndex !== undefined &&
      host?.type === "paragraph" &&
      isEmptyWatermarkHostParagraph(host)
    ) {
      next.content = [...next.content];
      next.content.splice(hostIndex, 1);
    }
    if (change.kind === "remove") {
      delete next.watermark;
      delete next.watermarkBlockIndex;
    } else {
      next.watermark = cloneModel(change.watermark);
      if (options.authority === "canonical" && !host) {
        const identity = options.hosts.find((entry) => entry.rId === rId);
        if (!identity)
          return refuse("An existing header needs an explicit watermark host identity.");
        const insertionIndex = Math.min(next.content.length, Math.max(0, hostIndex ?? 0));
        next.content = [...next.content];
        next.content.splice(insertionIndex, 0, {
          type: "paragraph",
          paraId: identity.paraId,
          content: [],
        });
        next.watermarkBlockIndex = insertionIndex;
      }
    }
    delete next.rawWatermarkXml;
    nextHeaders.set(rId, next);
  }
  const body = document.package.document;
  const nextBody =
    authority === "canonical" ? syncCanonicalSectionHeaders({ body, headers: nextHeaders }) : body;
  const updated = {
    ...document,
    package: { ...document.package, headers: nextHeaders, document: nextBody },
  };
  return change.kind === "remove"
    ? Result.ok(updated)
    : ensureDocumentWatermarkHeaderCoverage({
        document: updated,
        authority,
        watermark: change.watermark,
        coverage,
      });
};
export const applyDocumentWatermark = (document: Document, op: SetDocumentWatermarkOp) => {
  if (
    !Array.isArray(op.coverage) ||
    !Array.isArray(op.hosts) ||
    !isPayloadRecord(op.change) ||
    (op.change.kind !== "set" && op.change.kind !== "remove")
  )
    return refuse("The watermark operation payload is invalid.");
  const coverageIssue = identityPayloadIssue(op.coverage, "coverage");
  const hostsIssue = identityPayloadIssue(op.hosts, "hosts");
  if (coverageIssue || hostsIssue)
    return refuse(coverageIssue ?? hostsIssue ?? "Invalid watermark identities.");
  if (op.change.kind === "set") {
    const issue = watermarkPayloadIssue(op.change.watermark);
    if (issue) return refuse(issue);
  }
  const occupied = new Set(packageParagraphIds(document.package).map(idKey));
  for (const entry of [...op.coverage, ...op.hosts]) {
    if (typeof entry.rId !== "string" || typeof entry.paraId !== "string")
      return refuse("Watermark identities must be strings.");
    if (!isParaId(entry.paraId) || occupied.has(idKey(entry.paraId)))
      return refuse("Watermark coverage needs fresh package paragraph identities.");
    occupied.add(idKey(entry.paraId));
  }
  const neededHosts = op.change.kind === "set" ? planDocumentWatermarkHosts(document) : [];
  if (
    op.hosts.length !== neededHosts.length ||
    new Set(op.hosts.map(({ rId }) => rId)).size !== op.hosts.length ||
    neededHosts.some((rId) => !op.hosts.some((entry) => entry.rId === rId))
  )
    return refuse("Watermark hosts must name exactly the existing headers without a valid host.");
  if (op.change.kind === "set" && op.change.watermark.kind === "picture") {
    const watermark = op.change.watermark;
    if (typeof watermark.imageRId !== "string" || watermark.imageRId.trim() === "")
      return refuse("A picture watermark needs an image relationship identity.");
    if (watermark.imageTargetExternal) {
      const target = Result.try(() => new URL(watermark.imageTarget ?? ""));
      if (
        target.isErr() ||
        (target.value.protocol !== "https:" && target.value.protocol !== "http:")
      )
        return refuse("External watermark images require an HTTP or HTTPS target.");
    }
    if (
      !watermark.imageTargetExternal &&
      watermark.imageTarget !== undefined &&
      (watermark.imageTarget.startsWith("/") ||
        watermark.imageTarget.includes("..") ||
        /[:\\]/u.test(watermark.imageTarget))
    )
      return refuse("Embedded watermark targets must stay inside the package.");
  }
  if (op.change.kind === "set" && op.change.watermark.kind === "picture") {
    const normalized = normalizeCanonicalWatermark(op.change.watermark);
    if (
      normalized.kind === "picture" &&
      [normalized.widthPt, normalized.heightPt].some(
        (value) => value === undefined || !Number.isFinite(value) || /e/iu.test(String(value)),
      )
    )
      return refuse("Watermark dimensions must serialize as finite decimal lengths.");
  }
  const changed = setDocumentWatermarkWithCoverage({
    document,
    authority: "canonical",
    hosts: op.hosts,
    change:
      op.change.kind === "set"
        ? { kind: "set", watermark: normalizeCanonicalWatermark(op.change.watermark) }
        : { kind: "remove" },
    coverage: op.coverage.map(({ type, rId, paraId }) => ({
      type,
      rId,
      content: [{ type: "paragraph", paraId, content: [] }],
    })),
  });
  return changed.andThen((after) => storyLifecycleEdit(document, after, op));
};
