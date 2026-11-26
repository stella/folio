/** Shared helpers for tests that plan and apply small synthetic packages. */
import { Result } from "better-result";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import type {
  ApplyTemplatePlanResult,
  TemplatePlan,
  TemplatePlanError,
  TemplatePlanErrorCode,
  TemplatePlanItem,
  TemplatePlanReport,
} from "../types";
import { verifyTemplatePlanOutput } from "../verify";
import { buildMinimalDocx, type FixtureParts, NS, XML } from "./fixtures";

export const unwrap = <T>(result: Result<T, unknown>): T => {
  if (Result.isError(result)) {
    throw result.error;
  }
  return result.value;
};

export type PlannedBlock = { id: string; hash: string };

export type PlanOptions = {
  source: Uint8Array;
  template: Uint8Array;
  items?: (blocks: PlannedBlock[]) => TemplatePlanItem[];
  extra?: Partial<TemplatePlan>;
};

/** A plan that places every top-level block of the source, unless told otherwise. */
export const planFor = async ({
  source,
  template,
  items,
  extra = {},
}: PlanOptions): Promise<TemplatePlan> => {
  // A plan for one view of tracked changes names that view's blocks.
  const manifest = unwrap(
    await analyzeDocxBlocks(
      source,
      extra.trackedChanges === undefined
        ? {}
        : { trackedChanges: extra.trackedChanges },
    ),
  );
  const inspection = unwrap(await inspectDocxTemplate(template));
  const blocks = manifest.blocks.map(({ id, hash }) => ({ id, hash }));
  return {
    version: 1,
    sourceSha256: manifest.sha256,
    templateSha256: inspection.sha256,
    body:
      items?.(blocks) ??
      blocks.map(({ id, hash }) => ({ kind: "source", blockId: id, hash })),
    omit: [],
    ...extra,
  };
};

export type Applied = Result<ApplyTemplatePlanResult, TemplatePlanError>;

export const applyFor = async (options: PlanOptions): Promise<Applied> =>
  await applyTemplatePlan({
    source: options.source,
    template: options.template,
    plan: await planFor(options),
  });

/** The refusal code of an apply, or `"applied"`. */
export const codeOf = (result: Applied): TemplatePlanErrorCode | "applied" =>
  Result.isError(result) ? result.error.code : "applied";

/** The issue codes of the report a failed verification carries. */
export const issuesOf = (result: Applied): string[] =>
  Result.isError(result)
    ? (result.error.report?.issues.map(({ code }) => code) ?? [])
    : [];

export type VerifyOptions = {
  source: Uint8Array;
  template: Uint8Array;
  output: Uint8Array;
  plan: TemplatePlan;
};

export const verifyFor = async (
  options: VerifyOptions,
): Promise<TemplatePlanReport> =>
  unwrap(await verifyTemplatePlanOutput(options));

const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export type FixtureRelationship = {
  id: string;
  type: string;
  target: string;
  external?: boolean;
};

export const relationshipsXml = (
  items: readonly FixtureRelationship[],
): string =>
  `${XML}<Relationships xmlns="${REL}">${items
    .map(
      ({ id, type, target, external }) =>
        `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`,
    )
    .join("")}</Relationships>`;

export type SourceOptions = {
  body: string;
  relationships?: readonly FixtureRelationship[];
  /** Markup inside `w:footnotes`, after nothing: separators are not added. */
  footnotes?: string;
  styles?: string;
  numbering?: string;
  parts?: FixtureParts;
};

/** A source package around `body` with the parts a case needs. */
export const buildSourceWith = async ({
  body,
  relationships = [],
  footnotes,
  styles,
  numbering,
  parts = {},
}: SourceOptions): Promise<Uint8Array> => {
  const all: FixtureRelationship[] = [...relationships];
  const extra: FixtureParts = { ...parts };
  if (footnotes !== undefined) {
    all.push({ id: "rId90", type: "footnotes", target: "footnotes.xml" });
    extra["word/footnotes.xml"] =
      `${XML}<w:footnotes ${NS}>${footnotes}</w:footnotes>`;
  }
  if (styles !== undefined) {
    all.push({ id: "rId91", type: "styles", target: "styles.xml" });
    extra["word/styles.xml"] = `${XML}<w:styles ${NS}>${styles}</w:styles>`;
  }
  if (numbering !== undefined) {
    all.push({ id: "rId92", type: "numbering", target: "numbering.xml" });
    extra["word/numbering.xml"] =
      `${XML}<w:numbering ${NS}>${numbering}</w:numbering>`;
  }
  if (all.length > 0) {
    extra["word/_rels/document.xml.rels"] = relationshipsXml(all);
  }
  return await buildMinimalDocx(body, extra);
};

export const PIXEL_PNG: Uint8Array = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
    "base64",
  ),
);

/** A run holding an inline picture bound to `relationshipId`. */
export const pictureRun = (relationshipId: string, drawingId = 7): string =>
  `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="100000" cy="100000"/><wp:docPr id="${drawingId}" name="Picture ${drawingId}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const REVISION = 'w:author="a" w:date="2020-01-01T00:00:00Z"';

/** The attributes a tracked-change element carries. */
export const revision = (id: number): string => `w:id="${id}" ${REVISION}`;
