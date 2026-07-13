/**
 * Multi-engine DOCX comparison: static HTML report writer.
 *
 * Renders a `CorpusReport` (see `types.ts`) plus per-doc rendered assets into
 * a self-contained static site under `REPORT_DIR` (`config.ts`): an
 * `index.html` overview (per-doc scores, cross-corpus clusters) and one
 * `doc-<slug>.html` detail page per document (side-by-side reference/folio
 * page renderings with line-box overlays, plus the grouped divergence list).
 *
 * Deliberately imports only `./types` (and `./config` for the output path):
 * this module must stay testable standalone, without the renderer/folio/compare
 * extractors that produce its input.
 */
import { copyFile, mkdir, rm } from "node:fs/promises";
import path from "node:path";

import { REPORT_DIR } from "./config";
import type {
  Cluster,
  CorpusReport,
  Divergence,
  DivergenceKind,
  DocGeom,
  FeatureAttributedResult,
  LineBox,
  PageGeom,
} from "./types";

export type DocAssets = {
  referencePagePngs: string[];
  folioPagePngs: string[];
  referenceGeom: DocGeom;
  folioGeom: DocGeom;
};

const ASSETS_DIR_NAME = "assets";

const SCORE_GOOD = 0.98;
const SCORE_OK = 0.9;
const COLOR_GOOD = "#1a7f37";
const COLOR_OK = "#9a6700";
const COLOR_BAD = "#cf222e";

/** Report/detail-page divergence ordering, per the fixed contract. */
const DIVERGENCE_KIND_ORDER: DivergenceKind[] = [
  "page-count",
  "pagination",
  "missing-line",
  "extra-line",
  "line-break",
  "text-mismatch",
  "y-drift",
  "x-drift",
  "width-drift",
];

/** Per-doc page assets after copying into REPORT_DIR, relative-path or null (missing). */
type CopiedPageAssets = {
  referencePngs: (string | null)[];
  folioPngs: (string | null)[];
};

export const writeHtmlReport = async (
  report: CorpusReport,
  assets: Map<string, DocAssets>,
): Promise<string> => {
  await rm(REPORT_DIR, { recursive: true, force: true });
  await mkdir(REPORT_DIR, { recursive: true });
  await mkdir(path.join(REPORT_DIR, ASSETS_DIR_NAME), { recursive: true });

  const slugs = buildSlugs(report.results.map((result) => result.file));

  for (const result of report.results) {
    const slug = slugs.get(result.file);
    if (!slug) continue;

    const docAssets = assets.get(result.file);
    const copied = docAssets
      ? await copyDocAssets(slug, docAssets)
      : { referencePngs: [], folioPngs: [] };

    const html = renderDocPage(result, docAssets, copied, report.reference.displayName);
    await Bun.write(path.join(REPORT_DIR, `doc-${slug}.html`), html);
  }

  const indexPath = path.join(REPORT_DIR, "index.html");
  await Bun.write(indexPath, renderIndex(report, slugs));

  return indexPath;
};

// --- asset staging ---

const sanitizeSlug = (name: string): string => {
  const cleaned = name
    .replace(/[^a-zA-Z0-9-_]+/gu, "-")
    .replace(/^-+|-+$/gu, "")
    .toLowerCase();
  return cleaned.length > 0 ? cleaned : "doc";
};

const buildSlugs = (files: string[]): Map<string, string> => {
  const slugs = new Map<string, string>();
  const used = new Set<string>();

  for (const file of files) {
    const base = sanitizeSlug(path.basename(file, path.extname(file)));
    let candidate = base;
    let suffix = 2;
    while (used.has(candidate)) {
      candidate = `${base}-${suffix}`;
      suffix += 1;
    }
    used.add(candidate);
    slugs.set(file, candidate);
  }

  return slugs;
};

const copyDocAssets = async (slug: string, docAssets: DocAssets): Promise<CopiedPageAssets> => {
  const dir = path.join(REPORT_DIR, ASSETS_DIR_NAME, slug);
  await mkdir(dir, { recursive: true });

  const copyOne = async (srcPath: string, destName: string): Promise<string | null> => {
    try {
      await copyFile(srcPath, path.join(dir, destName));
      return `${ASSETS_DIR_NAME}/${slug}/${destName}`;
    } catch {
      return null;
    }
  };

  const referencePngs = await Promise.all(
    docAssets.referencePagePngs.map((src, i) => copyOne(src, `reference-${i + 1}.png`)),
  );
  const folioPngs = await Promise.all(
    docAssets.folioPagePngs.map((src, i) => copyOne(src, `folio-${i + 1}.png`)),
  );

  return { referencePngs, folioPngs };
};

// --- shared helpers ---

const escapeHtml = (value: string): string =>
  value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;")
    .replace(/'/gu, "&#39;");

const truncate = (text: string, max: number): string =>
  text.length > max ? `${text.slice(0, max - 1)}…` : text;

/** Table cell for a possibly-long text value: truncated, full text in `title`. */
const truncatedCell = (text: string, max: number): string => {
  if (text.length <= max) {
    return escapeHtml(text);
  }
  return `<span title="${escapeHtml(text)}">${escapeHtml(truncate(text, max))}</span>`;
};

const scoreColor = (score: number): string => {
  if (score >= SCORE_GOOD) return COLOR_GOOD;
  if (score >= SCORE_OK) return COLOR_OK;
  return COLOR_BAD;
};

/** Representative text for a divergence, independent of its kind's exact shape. */
const divergenceText = (divergence: Divergence): string => {
  switch (divergence.kind) {
    case "page-count":
      return `page count: reference=${divergence.reference}, folio=${divergence.folio}`;
    case "pagination":
      return divergence.text;
    case "line-break":
      return `${divergence.referenceTexts.join(" ")} | ${divergence.folioTexts.join(" ")}`;
    case "missing-line":
    case "extra-line":
    case "x-drift":
    case "y-drift":
    case "width-drift":
      return divergence.text;
    case "text-mismatch":
      return `${divergence.referenceText} → ${divergence.folioText}`;
    default: {
      const exhaustive: never = divergence;
      return exhaustive;
    }
  }
};

/** Page label for a divergence, independent of its kind's exact shape. */
const divergencePage = (divergence: Divergence): string => {
  switch (divergence.kind) {
    case "page-count":
      return "—";
    case "pagination":
      return `${divergence.referencePage} → ${divergence.folioPage}`;
    default:
      return String(divergence.page);
  }
};

/** Geometric delta for a divergence, when the kind carries one. */
const divergenceMagnitude = (divergence: Divergence): number | undefined => {
  switch (divergence.kind) {
    case "x-drift":
    case "width-drift":
      return divergence.deltaPt;
    case "y-drift":
      return divergence.residualPt;
    default:
      return undefined;
  }
};

const compactDivergenceCounts = (divergences: Divergence[]): string => {
  const counts = new Map<DivergenceKind, number>();
  for (const divergence of divergences) {
    counts.set(divergence.kind, (counts.get(divergence.kind) ?? 0) + 1);
  }
  return DIVERGENCE_KIND_ORDER.filter((kind) => (counts.get(kind) ?? 0) > 0)
    .map((kind) => `${counts.get(kind)} ${kind}`)
    .join(", ");
};

// --- index.html ---

const renderIndex = (report: CorpusReport, slugs: Map<string, string>): string => {
  const rows = report.results.map((result) => renderIndexRow(result, slugs)).join("\n");
  const clusterRows = report.clusters.map(renderClusterRow).join("\n");
  const docCount = report.results.length;
  const referenceVersion = report.reference.version ?? "unknown version";

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>folio DOCX interoperability report</title>
<style>${STYLE}</style>
</head>
<body>
<header>
<h1>DOCX rendering interoperability report</h1>
<p>Folio compared with ${escapeHtml(report.reference.displayName)} ${escapeHtml(referenceVersion)} &middot; generated ${escapeHtml(report.generatedAt)} &middot; ${docCount} document${docCount === 1 ? "" : "s"}</p>
<p class="method-note">The external renderer is a comparison reference, not a specification oracle. Agreement is interoperability evidence; OOXML conformance is assessed separately against the standard and structural validators.</p>
</header>
<main>
<section>
<h2>Documents</h2>
<table>
<thead><tr><th>Document</th><th>Score</th><th>Pages (reference / folio)</th><th>Lines matched</th><th>Median Y offset (pt)</th><th>Divergences</th></tr></thead>
<tbody>
${rows.length > 0 ? rows : `<tr><td colspan="6">No documents.</td></tr>`}
</tbody>
</table>
</section>
<section>
<h2>Clusters</h2>
<table>
<thead><tr><th>Kind</th><th>Feature</th><th>Count</th><th>Lift</th><th>Mean magnitude (pt)</th><th>Affected docs</th><th>Example</th></tr></thead>
<tbody>
${clusterRows.length > 0 ? clusterRows : `<tr><td colspan="7">No clusters.</td></tr>`}
</tbody>
</table>
</section>
</main>
</body>
</html>
`;
};

const renderIndexRow = (result: FeatureAttributedResult, slugs: Map<string, string>): string => {
  const slug = slugs.get(result.file) ?? "";
  const name = path.basename(result.file);
  const pct = (result.score * 100).toFixed(1);
  const counts = compactDivergenceCounts(result.divergences);

  return `<tr>
<td><a href="doc-${escapeHtml(slug)}.html">${escapeHtml(name)}</a></td>
<td><span class="score" style="color:${scoreColor(result.score)}">${pct}%</span></td>
<td>${result.referencePages} / ${result.folioPages}</td>
<td>${result.matchedLines} / ${result.totalReferenceLines}</td>
<td>${result.medianYOffsetPt.toFixed(2)}</td>
<td>${counts.length > 0 ? escapeHtml(counts) : "—"}</td>
</tr>`;
};

const renderClusterRow = (cluster: Cluster): string => {
  const example = cluster.examples.at(0);
  const exampleFullText = example ? divergenceText(example.divergence) : "";
  const docs = cluster.docs.map((doc) => path.basename(doc)).join(", ");

  return `<tr>
<td>${escapeHtml(cluster.kind)}</td>
<td>${escapeHtml(cluster.feature)}</td>
<td>${cluster.count}</td>
<td>${cluster.lift.toFixed(2)}</td>
<td>${cluster.meanMagnitudePt !== undefined ? cluster.meanMagnitudePt.toFixed(2) : "—"}</td>
<td>${escapeHtml(docs)}</td>
<td>${truncatedCell(exampleFullText, 80)}</td>
</tr>`;
};

// --- doc-<slug>.html ---

type PanelSide = "reference" | "folio";

const renderDocPage = (
  result: FeatureAttributedResult,
  docAssets: DocAssets | undefined,
  copied: CopiedPageAssets,
  referenceDisplayName: string,
): string => {
  const name = path.basename(result.file);
  const pageCount = docAssets
    ? Math.max(docAssets.referenceGeom.pages.length, docAssets.folioGeom.pages.length)
    : 0;

  const pagePairs = Array.from({ length: pageCount }, (_, i) =>
    renderPagePair(i, docAssets, copied, referenceDisplayName),
  ).join("\n");

  const banner =
    result.divergences.length === 0
      ? `<div class="banner-good">Full parity: no divergences detected.</div>`
      : "";

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(name)} — parity detail</title>
<style>${STYLE}</style>
</head>
<body>
<header>
<p><a href="index.html">&larr; back to summary</a></p>
<h1>${escapeHtml(name)}</h1>
<p>Folio compared with ${escapeHtml(referenceDisplayName)} &middot; score ${(result.score * 100).toFixed(1)}% &middot; ${result.matchedLines}/${result.totalReferenceLines} lines matched &middot; median Y offset ${result.medianYOffsetPt.toFixed(2)}pt</p>
</header>
<main>
${banner}
<section class="pages">
<h2>Pages</h2>
${pageCount > 0 ? pagePairs : `<p>No page assets available for this document.</p>`}
</section>
<section class="divergences">
<h2>Divergences</h2>
${renderDivergenceSections(result.divergences)}
</section>
</main>
<script>
document.querySelectorAll(".cross-toggle").forEach((el) => {
  el.addEventListener("change", () => {
    const target = document.getElementById(el.dataset.target);
    if (target) target.style.display = el.checked ? "block" : "none";
  });
});
</script>
</body>
</html>
`;
};

const renderPagePair = (
  pageIndex: number,
  docAssets: DocAssets | undefined,
  copied: CopiedPageAssets,
  referenceDisplayName: string,
): string => {
  const referencePage = docAssets?.referenceGeom.pages[pageIndex];
  const folioPage = docAssets?.folioGeom.pages[pageIndex];
  const referenceImg = copied.referencePngs[pageIndex] ?? null;
  const folioImg = copied.folioPngs[pageIndex] ?? null;

  const referencePanel = renderPanel({
    side: "reference",
    pageIndex,
    page: referencePage,
    otherPage: folioPage,
    imgPath: referenceImg,
    label: `${referenceDisplayName} — page ${pageIndex + 1}`,
  });
  const folioPanel = renderPanel({
    side: "folio",
    pageIndex,
    page: folioPage,
    otherPage: referencePage,
    imgPath: folioImg,
    label: `folio — page ${pageIndex + 1}`,
  });

  return `<div class="page-pair">${referencePanel}${folioPanel}</div>`;
};

/** Parameters for rendering one reference/folio page panel. */
type RenderPanelOptions = {
  side: PanelSide;
  pageIndex: number;
  page: PageGeom | undefined;
  otherPage: PageGeom | undefined;
  imgPath: string | null;
  label: string;
};

const renderPanel = ({
  side,
  pageIndex,
  page,
  otherPage,
  imgPath,
  label,
}: RenderPanelOptions): string => {
  if (!page) {
    const missingSide = side === "reference" ? "reference" : "folio";
    return `<div class="page-panel placeholder-panel">
<h3>${escapeHtml(label)}</h3>
<div class="placeholder">No corresponding ${missingSide} page</div>
</div>`;
  }

  const pageNum = pageIndex + 1;
  const crossId = `${side}-cross-${pageNum}`;
  const otherLabel = side === "reference" ? "folio" : "reference";
  const primaryClass = side === "reference" ? "reference-boxes" : "folio-boxes";
  const crossClass = side === "reference" ? "folio-boxes" : "reference-boxes";

  const primaryBoxes = renderBoxesSvgGroup(page.lines, primaryClass);
  const crossBoxes = otherPage
    ? renderBoxesSvgGroup(otherPage.lines, `${crossClass} cross`, crossId)
    : "";
  const toggle = otherPage
    ? `<label class="toggle"><input type="checkbox" class="cross-toggle" data-target="${escapeHtml(crossId)}" /> show ${escapeHtml(otherLabel)} boxes</label>`
    : "";
  const image = imgPath
    ? `<img src="${escapeHtml(imgPath)}" alt="${escapeHtml(label)}" />`
    : `<div class="placeholder img-placeholder" style="aspect-ratio: ${page.widthPt} / ${page.heightPt};">screenshot unavailable</div>`;

  return `<div class="page-panel">
<h3>${escapeHtml(label)}</h3>
${toggle}
<div class="page-image-wrap">
${image}
<svg viewBox="0 0 ${page.widthPt} ${page.heightPt}" preserveAspectRatio="none">
${primaryBoxes}
${crossBoxes}
</svg>
</div>
</div>`;
};

const renderBoxesSvgGroup = (lines: LineBox[], className: string, id?: string): string => {
  const idAttr = id ? ` id="${escapeHtml(id)}"` : "";
  const hiddenStyle = className.includes("cross") ? ' style="display:none"' : "";
  const rects = lines
    .map(
      (line) =>
        `<rect x="${line.xPt}" y="${line.yPt}" width="${line.widthPt}" height="${line.heightPt}" />`,
    )
    .join("");
  return `<g class="${escapeHtml(className)}"${idAttr}${hiddenStyle}>${rects}</g>`;
};

const renderDivergenceSections = (divergences: Divergence[]): string => {
  if (divergences.length === 0) {
    return "<p>None.</p>";
  }

  const byKind = new Map<DivergenceKind, Divergence[]>();
  for (const divergence of divergences) {
    const list = byKind.get(divergence.kind);
    if (list) {
      list.push(divergence);
    } else {
      byKind.set(divergence.kind, [divergence]);
    }
  }

  return DIVERGENCE_KIND_ORDER.filter((kind) => (byKind.get(kind)?.length ?? 0) > 0)
    .map((kind) => renderDivergenceKindSection(kind, byKind.get(kind) ?? []))
    .join("\n");
};

const renderDivergenceKindSection = (kind: DivergenceKind, divergences: Divergence[]): string => {
  const rows = divergences.map(renderDivergenceRow).join("\n");
  return `<h3>${escapeHtml(kind)} (${divergences.length})</h3>
<table>
<thead><tr><th>Page</th><th>Text</th><th>Delta</th></tr></thead>
<tbody>
${rows}
</tbody>
</table>`;
};

const renderDivergenceRow = (divergence: Divergence): string => {
  const page = divergencePage(divergence);
  const magnitude = divergenceMagnitude(divergence);
  const delta = magnitude !== undefined ? `${magnitude.toFixed(2)}pt` : "—";
  return `<tr><td>${escapeHtml(page)}</td><td>${truncatedCell(divergenceText(divergence), 120)}</td><td>${delta}</td></tr>`;
};

// --- shared inline CSS ---

const STYLE = `
:root { color-scheme: light; }
* { box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; margin: 0; background: #f6f8fa; color: #1f2328; }
header { padding: 24px 32px; background: #ffffff; border-bottom: 1px solid #d0d7de; }
header h1 { margin: 0 0 4px; font-size: 1.5rem; }
header p { margin: 0; color: #57606a; }
.method-note { margin-top: 8px; max-width: 90ch; }
main { padding: 24px 32px; }
section { margin-bottom: 32px; }
h2 { font-size: 1.15rem; border-bottom: 1px solid #d0d7de; padding-bottom: 6px; }
table { border-collapse: collapse; width: 100%; background: #fff; }
th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid #eaeef2; font-size: 0.9rem; vertical-align: top; }
th { background: #f6f8fa; font-weight: 600; }
.score { font-weight: 700; }
a { color: #0969da; text-decoration: none; }
a:hover { text-decoration: underline; }
.banner-good { background: #dafbe1; border: 1px solid #1a7f37; color: #1a7f37; padding: 12px 16px; border-radius: 6px; margin-bottom: 16px; font-weight: 600; }
.page-pair { display: flex; gap: 16px; margin-bottom: 24px; align-items: flex-start; }
.page-panel { flex: 1 1 0; min-width: 0; background: #fff; border: 1px solid #d0d7de; border-radius: 6px; padding: 12px; }
.page-panel h3 { margin: 0 0 8px; font-size: 0.95rem; }
.toggle { display: block; font-size: 0.8rem; margin-bottom: 8px; color: #57606a; }
.page-image-wrap { position: relative; width: 100%; background: #eaeef2; line-height: 0; }
.page-image-wrap img { display: block; width: 100%; height: auto; }
.page-image-wrap svg { position: absolute; top: 0; left: 0; width: 100%; height: 100%; }
.reference-boxes rect { stroke: #0969da; fill: none; stroke-width: 1; vector-effect: non-scaling-stroke; }
.folio-boxes rect { stroke: #cf222e; fill: none; stroke-width: 1; vector-effect: non-scaling-stroke; }
.placeholder { display: flex; align-items: center; justify-content: center; background: #eaeef2; color: #57606a; border: 1px dashed #9198a1; min-height: 160px; font-size: 0.85rem; }
.placeholder-panel { flex: 1 1 0; }
`;
