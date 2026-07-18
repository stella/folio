# Plan: Upstream docx-editor investigation and resync

Date: 2026-07-18

## Goal

Record what happened to `eigenpal/docx-editor` during the June–July 2026
outage, what changed while it was dark / immediately on return, and which
upstream deltas are worth porting into folio. This is an internal investigation
artifact; no product README or public messaging changes.

## What likely happened

Upstream was publicly unreachable from roughly **2026-06-26** (HN report:
[item 48692474](https://news.ycombinator.com/item?id=48692474)) through about
**2026-07-16/17**. Wayback CDX shows `github.com/eigenpal/docx-editor` as `404`
on 2026-07-09 and `200` again on 2026-07-17; `docx-editor.dev` returned `503`
on 2026-07-04 and recovered earlier than GitHub.

No public maintainer post-mortem was found. The strongest *internal* evidence
from their commit/PR stream is a **security-driven withdrawal**, not a product
abandonment:

1. **Immediate pre-outage security blitz (2026-06-20 → 06-22)**
   - `#952` untrusted-DOCX security checklist in `CLAUDE.md`
   - `#953` `SECURITY.md`
   - `#960` clipboard HTML XSS + ReDoS (DOMPurify; stop live `innerHTML` parse)
   - `#961`/`#962` polynomial ReDoS / font-name escape / data-URL regex
   - `#982` zip-bomb limits, image `src` scheme guard, MCP buffer caps,
     prototype-pollution guards
   - CI least-privilege (`#964`)
2. **Last `@eigenpal/*` npm release:** `1.9.0` on **2026-06-21**
3. **Repo inaccessible** while private work continued (git history is continuous;
   ~208 commits 2026-06-24 → 07-16, mostly fidelity `fix`es plus a rendering
   model rename)
4. **Return paired with package rebrand (2026-07-16 → 07-17)**
   - `#1065` publish under `@docx-editor.dev/{core,react,vue,agents,i18n,nuxt}`
   - `#1070` reset version baseline to `0.0.1` so the first real release is `1.0.0`
   - `#1071` scrub leftover `@eigenpal` references
   - `#1072` restore corrupted README PNG blobs + mark PNGs binary
   - `#1073` further clipboard / locale hardening (CodeQL)
   - npm currently only has `@docx-editor.dev/core` **placeholders**
     (`0.0.1`, `0.0.1-placeholder`); real code still ships as
     `@eigenpal/docx-editor-core@1.9.0`

Secondary signals (fixture ZIP central-directory truncation in `#1065`, PNG
line-ending corruption in `#1072`) look like **repo hygiene fallout** from the
private/restore process, not the root cause.

HN speculation (Microsoft IP / AI regurgitation) has **no corroboration** in
the public commit stream. Treat that as rumor.

Star-count collapse (~1200 → ~20 on restore) is consistent with
delete-and-repush or equivalent public-identity reset; history itself was
preserved.

## What changed upstream (relative to folio's last useful sync)

Folio already cites many eigenpal issues through roughly **`#994`**. The
meaningful gap is:

| Window | Upstream focus |
| --- | --- |
| Pre-outage, unported | Security `#960`–`#982`; text-box/bookmark round-trip `#967`; Word paste `#984`; EMF/WMF HF + smartTag `#985`; options-aware TOC `#987`; HF band per section `#988`; live editable footnotes `#995` |
| Private window | Large rendering rename (`flow-model` / `pagination-model` / `painter-model`); hundreds of Word-fidelity pagination/HF/table/float/tab fixes; TOC regeneration stack |
| On return | npm scope rebrand; structural revision margin indicators `#1068`; deleted section-break pagination `#1067`; tab-grid / border-flow / floating-table `#1069`; embedded-font metric sharing `#1064`; inline TOC refresh `#1074` |

Architectures have diverged. Upstream renamed layout packages; folio kept
`layout-bridge` / `layout-engine` / `layout-painter` / `paged-layout` /
`render-dom` and has grown past upstream in several areas (folio-core ~827 TS
files vs upstream-core ~647). **Do not merge trees.** Port invariants and
tests surgically.

## Design Decisions

- **Investigate first, port second:** This plan is the triage map; each port
  lands as its own PR with a synthetic regression test (fidelity consolidation
  rules).
- **Prefer OOXML/layout invariants over corpus docs:** Never branch on their
  "Comprehensive Word Element Test" fixture identity.
- **Security ports are independent of fidelity ports:** Ship security deltas
  even when layout ports are deferred.
- **Keep `@eigenpal/docx-editor-core@1.9.0` as the benchmark pin until
  `@docx-editor.dev/core` publishes a real non-placeholder release;** then
  switch benchmarks/migrate scripts deliberately.

## Scope

**In scope (recommended port queue):**

### P0 — Security / untrusted input

Already partially covered by folio `#420` (DoS, URL/CSS/OOXML injection,
agent scoping, hidden-content leak). Still review and close gaps against:

1. Clipboard HTML trust boundary (`#960`, follow-up `#1073`): DOMPurify + inert
   document walk; linear comment strip; XML-declaration scanner.
2. Zip / image / MCP guards from `#982` if any path is still weaker than
   folio's `DocxUnzipLimits`.
3. ReDoS / font-name escaping from `#961`/`#962` where folio still builds
   regexes from document or toolbar strings.

### P1 — User-visible features folio lacks or only partially has

1. **TOC regeneration** (`#987`, `#1058`, `#1074`): folio has `generateTOC`
   insert; missing stale detection, inline refresh affordance, bookmark-stable
   regenerate, options (level range / title / hyperlinks).
2. **Live footnote editing parity** (`#995`): folio has note save + some story
   host wiring; confirm click→edit→tracked-change parity with upstream.
3. **Word-style structural revision indicators** (`#1068`): margin bars /
   paragraph marks for structural insert/delete across body, tables, HF, images.
4. **Deleted section breaks stay reviewable without paginating** (`#1067`).

### P2 — Layout fidelity invariants (port as reusable rules)

From `#1069` and the private-window layout cluster:

1. Automatic tab-stop grid only resumes **past the rightmost custom `w:tabs`
   stop** (OOXML §17.15.1.25); right-tab PAGE/NUMPAGES spacing must not
   collapse.
2. Paragraph border + `w:space` consume flow height (adjacent bordered callouts
   must not touch).
3. Floating / positioned table anchoring vs Word.
4. Embedded-font metrics shared across duplicate bundles (`#1064`).
5. HF band growth per section margins (`#988`).
6. Text-box + block-level bookmark round-trip (`#967`).
7. Keep footnote ref on same line as its word (`#991`); superscript only when
   style says so (`#994`) — verify folio already matches.

### P3 — Housekeeping (internal only)

1. Track upstream package rename for benchmarks / migrate codemod when real
   `@docx-editor.dev/*` versions ship.
2. Optionally mirror their SECURITY.md process locally if folio still lacks one
   (separate decision; not part of fidelity work).

**Out of scope:**

- Rewriting folio onto upstream's `flow-model` / `painter-model` rename.
- Porting Vue-only e2e churn, demo fixture defaults (`#1079`), or npm scope
  cosmetics.
- Public README / marketing updates about upstream being back.
- Blind cherry-picks of multi-hundred-commit private-window history.

## Implementation

No code in this PR. Follow-up PRs should each:

1. Name the upstream PR/commit and the invariant.
2. Map to folio modules (`packages/core/src/docx/*`, `layout-bridge/*`,
   `layout-engine/*`, `layout-painter/*`, `prosemirror/*`, clipboard utils).
3. Add a minimal synthetic test; keep their corpus docs out of the tree.
4. Add a changeset when `packages/*/src` changes.

Suggested first follow-ups (separate PRs):

1. `fix(security): align clipboard HTML paste with upstream #960/#1073`
2. `feat(core): TOC stale detection + regenerate (upstream #1058/#1074)`
3. `fix(core): ignore tracked-deleted section breaks in pagination (#1067)`
4. `fix(core): tab-stop grid + bordered-paragraph flow height (#1069 subset)`
5. `feat(core): Word-style structural revision indicators (#1068)`

## Test Cases

For each ported invariant:

- Unit test on synthetic OOXML / PM doc proving the rule.
- Where UI is involved (TOC refresh, revision bars, footnote edit): one
  playground interaction or parity e2e, React-first; Vue only if the seam is
  shared incorrectly.
- Security ports: adversarial HTML / ZIP / regex timing assertions (finish
  near-instantly).

## Open Questions

1. Confirm with maintainers (off-repo) whether the outage was security-only or
   also included legal/hosting constraints; do not guess in public artifacts.
2. When will `@docx-editor.dev/core` publish a real build we can pin in
   `benchmarks/*` instead of `@eigenpal/docx-editor-core@1.9.0`?
3. How much of `#995` live footnote editing does folio already cover via the
   Vue/React note story hosts?
4. Should folio treat upstream as an ongoing sync source again, or only
   cherry-pick high-value invariants?

## Evidence snapshot (checked 2026-07-18)

- Upstream HEAD: `8b708763` — `feat: add inline TOC regeneration (#1074)`
- Upstream homepage: `https://docx-editor.dev`
- License on upstream: Apache-2.0 (unchanged from folio's NOTICE lineage)
- Folio still depends on `@eigenpal/docx-editor-core@^1.9.0` for
  parse/serialize benchmarks and migrate tests
