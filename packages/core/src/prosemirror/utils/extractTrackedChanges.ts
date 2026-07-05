/**
 * Walk the PM doc once and derive (a) the tracked-change list and (b) a
 * comment→revision overlap map for threading. Adjacent entries from the
 * same revision are merged; deletion+insertion pairs from the same
 * author/date become a single `replacement` entry (matches Word's UX
 * for replace ops).
 *
 * Pure function — no React, no Vue, no side effects. Single O(N) walk
 * over text nodes.
 *
 * @packageDocumentation
 * @public
 */
import type { EditorState } from "prosemirror-state";
import type { Mark } from "prosemirror-model";

/**
 * One tracked change surfaced by {@link extractTrackedChanges}. Each entry
 * groups all sites of one revision into a single row that the sidebar
 * renders as one card.
 *
 * @public
 */
export interface TrackedChangeEntry {
  /**
   * Revision shape. Inline shapes (`insertion`, `deletion`, `replacement`)
   * wrap text runs; the rest are structural revisions on node attrs.
   *
   * - `insertion` — text was added (`<w:ins>`).
   * - `deletion` — text was struck through but not removed (`<w:del>`).
   * - `replacement` — a deletion + insertion by the same author at the
   *   same position+time; sidebar shows one combined card. `deletedText`
   *   and `insertionRevisionId` are set on this variant.
   * - `paragraphMarkInsertion` / `paragraphMarkDeletion` — Enter /
   *   Backspace produced a tracked paragraph break (`<w:pPr><w:rPr><w:ins/>` /
   *   `<w:del/>`).
   * - `paragraphPropertiesChanged` — formatting (alignment, spacing,
   *   etc.) on the paragraph was changed (`<w:pPrChange>`).
   * - `rowInserted` / `rowDeleted` / `rowPropertiesChanged` — table
   *   row authored / removed / formatted (`<w:trPr><w:ins/>` / `<w:del/>`
   *   / `<w:trPrChange>`).
   * - `cellInserted` / `cellDeleted` / `cellMerged` /
   *   `cellPropertiesChanged` — per-cell revisions
   *   (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>` / `<w:tcPrChange>`).
   * - `tablePropertiesChanged` — table-level formatting
   *   (`<w:tblPrChange>`).
   */
  type:
    | "insertion"
    | "deletion"
    | "replacement"
    | "paragraphMarkInsertion"
    | "paragraphMarkDeletion"
    | "paragraphPropertiesChanged"
    | "rowInserted"
    | "rowDeleted"
    | "rowPropertiesChanged"
    | "cellInserted"
    | "cellDeleted"
    | "cellMerged"
    | "cellPropertiesChanged"
    | "tableInserted"
    | "tableDeleted"
    | "tablePropertiesChanged";
  /**
   * Affected text. For inline types this is the run's text; for
   * structural types it's the surrounding paragraph / cell content
   * (truncated by the sidebar before display).
   */
  text: string;
  /**
   * Only set when `type === 'replacement'` — the text the user removed.
   * The inserted text lives in {@link TrackedChangeEntry.text}.
   */
  deletedText?: string | undefined;
  /** Author that minted the revision (`w:author`). */
  author: string;
  /** ISO timestamp the revision was minted (`w:date`). May be undefined for legacy imports. */
  date?: string | undefined;
  /**
   * Document position where the revision starts. For inline types this
   * is the start of the marked text run; for structural types it's the
   * containing paragraph / row / cell / table node's start position.
   */
  from: number;
  /**
   * Document position where the revision ends. For inline coalesced
   * runs that span multiple paragraphs, this is the END position of the
   * LAST run in the group; the intervening structural positions are not
   * preserved.
   */
  to: number;
  /** The `w:id` of the revision. */
  revisionId: number;
  /**
   * Only set when `type === 'replacement'` — the insertion half carries
   * a DIFFERENT `w:id` from the deletion (sharing would trip the OOXML
   * move-pair serializer).
   */
  insertionRevisionId?: number | undefined;
  /**
   * Extra `w:id`s that map to the same logical revision as this card.
   * Populated when the extractor coalesces a burst of distinct ids by
   * (author, date) — e.g. a foreign document where the source editor
   * minted a fresh id per atomic edit.
   */
  coalescedRevisionIds?: number[] | undefined;
}

/**
 * Output of {@link extractTrackedChanges}.
 *
 * @public
 */
export interface TrackedChangesResult {
  /** Tracked-change entries, sorted by document position, with adjacent same-revision entries merged. */
  entries: TrackedChangeEntry[];
  /**
   * Map of `commentId -> revisionId` for comments whose range overlaps a tracked-change mark.
   * Consumers (DocxEditor's threading effect) use this to thread comments under their tracked change.
   */
  commentToRevision: Map<number, number>;
}

const EMPTY_RESULT: TrackedChangesResult = {
  entries: [],
  commentToRevision: new Map(),
};

/**
 * Walk the PM doc and extract every tracked change as a flat list of
 * `TrackedChangeEntry` plus a comment→revision overlap map. Adjacent
 * inline marks coalesce by `(type, revisionId, author, date)`; a
 * deletion immediately followed by an insertion (same author + same
 * date) collapses into a single `replacement` entry; paragraph-mark
 * cards (`paragraphMarkInsertion` / `paragraphMarkDeletion`) are
 * hidden when an inline entry already covers their revision triple
 * (one Accept clears every site of one conceptual change).
 *
 * Pure and deterministic. Returns `EMPTY_RESULT` on null state.
 *
 * @public
 */
export function extractTrackedChanges(state: EditorState | null): TrackedChangesResult {
  if (!state) return EMPTY_RESULT;
  const { doc, schema } = state;
  const insertionType = schema.marks["insertion"];
  const deletionType = schema.marks["deletion"];
  const commentType = schema.marks["comment"];
  if (!insertionType && !deletionType) return EMPTY_RESULT;

  const raw: TrackedChangeEntry[] = [];
  const commentToRevision = new Map<number, number>();
  const commentMetadata = new Map<number, { revisions: Set<number>; hasCleanText: boolean }>();
  doc.descendants((node, pos) => {
    // Structural revisions on the paragraph mark itself
    // (`<w:pPr><w:rPr><w:ins/>` / `<w:del/>`). Surface as their own entry
    // types so the sidebar can label and dispatch them correctly.
    if (node.type.name === "paragraph") {
      const ins = node.attrs["pPrIns"] as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      const del = node.attrs["pPrDel"] as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      if (ins) {
        raw.push({
          type: "paragraphMarkInsertion",
          text: node.textContent || "",
          author: ins.author || "",
          date: ins.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: ins.revisionId,
        });
      }
      if (del) {
        raw.push({
          type: "paragraphMarkDeletion",
          text: node.textContent || "",
          author: del.author || "",
          date: del.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: del.revisionId,
        });
      }
      // Paragraph-property changes — one entry per (id, author, date) entry
      // in the pPrChange array. Reject restores prior values; accept clears.
      const pPrChange = node.attrs["pPrChange"] as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(pPrChange)) {
        for (const entry of pPrChange) {
          raw.push({
            type: "paragraphPropertiesChanged",
            text: node.textContent || "",
            author: entry.info.author || "",
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
      // Descend into paragraph content; do not return here.
    }

    // Table-row revisions (`<w:trPr><w:ins/>` / `<w:del/>` / `<w:trPrChange>`).
    if (node.type.name === "tableRow") {
      const trIns = node.attrs["trIns"] as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      const trDel = node.attrs["trDel"] as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      if (trIns) {
        raw.push({
          type: "rowInserted",
          text: node.textContent || "",
          author: trIns.author || "",
          date: trIns.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: trIns.revisionId,
        });
      }
      if (trDel) {
        raw.push({
          type: "rowDeleted",
          text: node.textContent || "",
          author: trDel.author || "",
          date: trDel.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: trDel.revisionId,
        });
      }
      const trPrChange = node.attrs["trPrChange"] as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(trPrChange)) {
        for (const entry of trPrChange) {
          if (!entry?.info || typeof entry.info.id !== "number") continue;
          raw.push({
            type: "rowPropertiesChanged",
            text: node.textContent || "",
            author: entry.info.author || "",
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
      // Descend into cells.
    }

    // Table-cell revisions (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>`,
    // `<w:tcPrChange>`).
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const cellMarker = node.attrs["cellMarker"] as {
        kind: "ins" | "del" | "merge";
        info: { revisionId: number; author: string; date: string | null };
      } | null;
      if (cellMarker?.info && typeof cellMarker.info.revisionId === "number") {
        const kindToType = {
          ins: "cellInserted" as const,
          del: "cellDeleted" as const,
          merge: "cellMerged" as const,
        };
        const resolvedType = kindToType[cellMarker.kind];
        if (resolvedType) {
          raw.push({
            type: resolvedType,
            text: node.textContent || "",
            author: cellMarker.info.author || "",
            date: cellMarker.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: cellMarker.info.revisionId,
          });
        }
      }
      const tcPrChange = node.attrs["tcPrChange"] as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(tcPrChange)) {
        for (const entry of tcPrChange) {
          if (!entry?.info || typeof entry.info.id !== "number") continue;
          raw.push({
            type: "cellPropertiesChanged",
            text: node.textContent || "",
            author: entry.info.author || "",
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
    }

    // Table-level property change (`<w:tblPrChange>`).
    if (node.type.name === "table") {
      const tblPrChange = node.attrs["tblPrChange"] as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(tblPrChange)) {
        for (const entry of tblPrChange) {
          if (!entry?.info || typeof entry.info.id !== "number") continue;
          raw.push({
            type: "tablePropertiesChanged",
            text: "",
            author: entry.info.author || "",
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }

      // Whole-table insertion / deletion: when every row carries a trIns
      // (or trDel) from the SAME (author, date) — not necessarily the same
      // `w:id`, since foreign editors mint a fresh id per row — surface ONE
      // `tableInserted` / `tableDeleted` entry. The per-row revision ids
      // get tucked into `coalescedRevisionIds` so accepting the card clears
      // every row's marker in one go.
      const firstRow = node.firstChild;
      const firstIns = firstRow?.attrs["trIns"] as
        | {
            revisionId: number;
            author?: string;
            date?: string | null;
          }
        | null
        | undefined;
      const firstDel = firstRow?.attrs["trDel"] as
        | {
            revisionId: number;
            author?: string;
            date?: string | null;
          }
        | null
        | undefined;
      const sharedAttr = firstIns ? "trIns" : firstDel ? "trDel" : null;
      if (sharedAttr) {
        const sharedRev = (firstIns ?? firstDel) as {
          revisionId: number;
          author: string;
          date: string | null;
        };
        let allShare = true;
        const rowRevIds: number[] = [];
        node.forEach((row) => {
          if (row.type.name !== "tableRow") {
            allShare = false;
            return;
          }
          const v = row.attrs[sharedAttr] as
            | { revisionId: number; author?: string; date?: string | null }
            | null
            | undefined;
          if (
            !v ||
            (v.author ?? "") !== (sharedRev.author ?? "") ||
            (v.date ?? null) !== (sharedRev.date ?? null)
          ) {
            allShare = false;
            return;
          }
          rowRevIds.push(v.revisionId);
        });
        if (allShare) {
          // Exclude text inside deletion marks: the empty-vs-content switch
          // downstream compares `text.trim().length` to decide whether to
          // surface "Inserted table" or defer to an inline card. Deletion-
          // marked text is still rendered in the doc but represents removed
          // content, so it shouldn't count as "the table has content".
          let visibleText = "";
          if (deletionType) {
            node.descendants((child) => {
              if (child.isText && !child.marks.some((m) => m.type === deletionType)) {
                visibleText += child.text || "";
              }
            });
          } else {
            visibleText = node.textContent || "";
          }
          const primaryId = sharedRev.revisionId;
          const extraIds = rowRevIds.filter(
            (id, idx) => id !== primaryId && rowRevIds.indexOf(id) === idx,
          );
          raw.push({
            type: sharedAttr === "trIns" ? "tableInserted" : "tableDeleted",
            text: visibleText,
            author: sharedRev.author || "",
            date: sharedRev.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: primaryId,
            ...(extraIds.length > 0 ? { coalescedRevisionIds: extraIds } : {}),
          });
        }
      }
    }

    // Text AND inline atoms (image, shape) can carry tracked-change marks, so
    // an inserted picture shows up in the sidebar like inserted text. Atoms
    // have no `.text`, so label them by node type (alt text when present).
    if (!node.isInline) return;
    const inlineText = node.isText
      ? node.text || ""
      : node.type.name === "image"
        ? (node.attrs["alt"] as string) || "image"
        : node.type.name;
    let tcMark: Mark | null = null;
    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        raw.push({
          type: mark.type === insertionType ? "insertion" : "deletion",
          text: inlineText,
          author: (mark.attrs["author"] as string) || "",
          date: mark.attrs["date"] as string | undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: mark.attrs["revisionId"] as number,
        });
        tcMark = mark;
      }
    }
    if (commentType) {
      const commentMark = node.marks.find((m) => m.type === commentType);
      if (commentMark) {
        const cid = commentMark.attrs["commentId"] as number;
        let meta = commentMetadata.get(cid);
        if (!meta) {
          meta = { revisions: new Set(), hasCleanText: false };
          commentMetadata.set(cid, meta);
        }
        if (tcMark) {
          meta.revisions.add(tcMark.attrs["revisionId"] as number);
        } else {
          meta.hasCleanText = true;
        }
      }
    }
  });

  for (const [cid, meta] of commentMetadata.entries()) {
    if (!meta.hasCleanText && meta.revisions.size === 1) {
      const rid = meta.revisions.values().next().value!;
      commentToRevision.set(cid, rid);
    }
  }

  // Coalesce structural-revision entries that share a `(id, author, date)`
  // triple across nested nodes. A row-insertion typically produces one
  // `rowInserted` entry on the `<tr>` PLUS one `cellInserted` entry per
  // cell, all sharing the triple. The spec says these should render as a
  // single sidebar row. Prefer the broader entry: priority is
  // `table > row > cell > paragraph-mark`.
  //
  // Inline insertion/deletion entries are NOT coalesced here — the
  // adjacent-merge pass below handles them.
  //
  // Single-pass: track each triple's slot index in `ordered` so an
  // in-place replacement is O(1) (vs `ordered.indexOf(existing)` which
  // would be O(n) inside an O(n) loop).
  const STRUCTURAL_PRIORITY: Record<string, number> = {
    tableInserted: 6,
    tableDeleted: 6,
    tablePropertiesChanged: 5,
    rowInserted: 4,
    rowDeleted: 4,
    rowPropertiesChanged: 4,
    cellInserted: 3,
    cellDeleted: 3,
    cellMerged: 3,
    cellPropertiesChanged: 3,
    paragraphMarkInsertion: 2,
    paragraphMarkDeletion: 2,
    paragraphPropertiesChanged: 2,
  };
  const isStructuralType = (t: TrackedChangeEntry["type"]) => t in STRUCTURAL_PRIORITY;
  const slotByKey = new Map<string, number>();
  const ordered: TrackedChangeEntry[] = [];
  // Helper: collect every distinct `w:id` involved in coalescing the dropped
  // entry into the survivor, EXCLUDING the survivor's own primary id.
  const mergeIds = (survivor: TrackedChangeEntry, dropped: TrackedChangeEntry): number[] => {
    const ids = new Set<number>(survivor.coalescedRevisionIds ?? []);
    for (const id of dropped.coalescedRevisionIds ?? []) ids.add(id);
    ids.add(dropped.revisionId);
    ids.delete(survivor.revisionId);
    return [...ids];
  };
  for (const entry of raw) {
    if (!isStructuralType(entry.type)) {
      ordered.push(entry);
      continue;
    }
    // Key by (author, date) only — foreign editors mint a fresh `w:id` per
    // atomic edit, so triples that differ only in id are still one logical
    // revision burst and should share a sidebar card. Same-author bursts
    // at distinct ms-precision timestamps stay separate.
    const key = `${entry.author}|${entry.date ?? ""}`;
    const slot = slotByKey.get(key);
    if (slot === undefined) {
      slotByKey.set(key, ordered.push(entry) - 1);
      continue;
    }
    const existing = ordered[slot]!;
    const incomingPri = STRUCTURAL_PRIORITY[entry.type] ?? 0;
    const existingPri = STRUCTURAL_PRIORITY[existing.type] ?? 0;
    if (incomingPri > existingPri) {
      // Incoming wins (broader scope). Carry the existing id forward.
      ordered[slot] = { ...entry, coalescedRevisionIds: mergeIds(entry, existing) };
    } else {
      // Existing stays; absorb the dropped id so accept clears every site.
      ordered[slot] = { ...existing, coalescedRevisionIds: mergeIds(existing, entry) };
    }
  }

  // Merge inline insertion/deletion entries that share a logical revision
  // burst (same type, author, date) into a single sidebar card. The
  // suggesting-mode plugin coalesces a continuous editing run under one
  // revisionId — including runs split across paragraph boundaries — but
  // foreign editors mint a fresh id per atomic edit, so dropping the
  // revisionId from the key keeps both cases consistent.
  const inlineGroups = new Map<string, TrackedChangeEntry>();
  const merged: TrackedChangeEntry[] = [];
  for (const entry of ordered) {
    const isInlineType = entry.type === "insertion" || entry.type === "deletion";
    if (!isInlineType) {
      merged.push({ ...entry });
      continue;
    }
    const key = `${entry.type}|${entry.author}|${entry.date ?? ""}`;
    const group = inlineGroups.get(key);
    if (group) {
      // Cross-paragraph runs get a space separator; literally adjacent runs
      // concatenate directly.
      const sep = group.to === entry.from ? "" : " ";
      group.text += sep + entry.text;
      group.to = entry.to;
      if (entry.revisionId !== group.revisionId) {
        const ids = new Set<number>(group.coalescedRevisionIds ?? []);
        for (const id of entry.coalescedRevisionIds ?? []) ids.add(id);
        ids.add(entry.revisionId);
        ids.delete(group.revisionId);
        group.coalescedRevisionIds = ids.size > 0 ? [...ids] : undefined;
      }
    } else {
      const copy = { ...entry };
      inlineGroups.set(key, copy);
      merged.push(copy);
    }
  }

  // Detect replacement pairs: adjacent deletion + insertion from the
  // same author/date. Word assigns different w:id values but same
  // author+date for a single replace.
  const final: TrackedChangeEntry[] = [];
  for (let i = 0; i < merged.length; i++) {
    const curr = merged[i]!;
    const next = merged[i + 1];
    if (
      curr.type === "deletion" &&
      next &&
      next.type === "insertion" &&
      curr.author === next.author &&
      curr.date === next.date &&
      curr.to === next.from
    ) {
      final.push({
        type: "replacement",
        text: next.text,
        deletedText: curr.text,
        author: curr.author,
        date: curr.date,
        from: curr.from,
        to: next.to,
        revisionId: curr.revisionId,
        insertionRevisionId: next.revisionId,
      });
      i++;
    } else {
      final.push(curr);
    }
  }

  // Final pass: if a paragraph-mark entry (pPrIns / pPrDel) shares its
  // (author, date) with an inline entry (insertion / deletion / replacement),
  // the inline entry already represents the whole conceptual edit — hide
  // the structural sibling so the sidebar shows ONE card per change.
  // Migrate the hidden entry's `coalescedRevisionIds` (and primary id) to
  // the surviving inline entry so one Accept still clears every site.
  const inlineByKey = new Map<string, TrackedChangeEntry>();
  for (const e of final) {
    if (e.type === "insertion" || e.type === "deletion" || e.type === "replacement") {
      const k = `${e.author}|${e.date ?? ""}`;
      if (!inlineByKey.has(k)) inlineByKey.set(k, e);
    }
  }
  for (const e of final) {
    if (e.type !== "paragraphMarkInsertion" && e.type !== "paragraphMarkDeletion") continue;
    const survivor = inlineByKey.get(`${e.author}|${e.date ?? ""}`);
    if (!survivor) continue;
    const ids = new Set<number>(survivor.coalescedRevisionIds ?? []);
    for (const id of e.coalescedRevisionIds ?? []) ids.add(id);
    ids.add(e.revisionId);
    ids.delete(survivor.revisionId);
    if (survivor.type === "replacement" && survivor.insertionRevisionId != null) {
      ids.delete(survivor.insertionRevisionId);
    }
    survivor.coalescedRevisionIds = ids.size > 0 ? [...ids] : undefined;
  }
  // Table-vs-inline preference:
  //   - Empty inserted/deleted table (no typed cell content) → show the
  //     `tableInserted`/`tableDeleted` card. There's no inline content
  //     to anchor the change to, so the structural label is the only
  //     thing the user can act on.
  //   - Table WITH content (user typed in cells after insert) → show the
  //     inline "Added 'X'" / "Replaced ..." card instead. The text is
  //     more informative than the generic "Inserted table" label, and
  //     one Accept on the inline card clears every site (cells share the
  //     same revisionId via the cellMarker-inherits-id rule).
  //
  // `tableKeys` collects only EMPTY table-level entries; the dedup pass
  // hides inline cards that share those keys. Non-empty table entries
  // are themselves dropped so the user sees the inline card.
  const emptyTableEntries = new Set<TrackedChangeEntry>();
  const tableByKey = new Map<string, TrackedChangeEntry>();
  for (const e of final) {
    if (e.type !== "tableInserted" && e.type !== "tableDeleted") continue;
    const key = `${e.author}|${e.date ?? ""}`;
    const hasContent = e.text.trim().length > 0;
    if (hasContent) {
      // drop this entry so the inline card represents the change
      continue;
    }
    emptyTableEntries.add(e);
    if (!tableByKey.has(key)) tableByKey.set(key, e);
  }
  // Mirror the inline-survivor migration above: when a paragraph-mark entry
  // is about to be hidden because it shares (author, date) with a surviving
  // empty-table entry, fold its `revisionId` (+ its own coalesced ids) into
  // the table's `coalescedRevisionIds`. Without this, accepting the table
  // card would leave the orphaned pPrIns/pPrDel behind.
  for (const e of final) {
    if (e.type !== "paragraphMarkInsertion" && e.type !== "paragraphMarkDeletion") continue;
    const survivor = tableByKey.get(`${e.author}|${e.date ?? ""}`);
    if (!survivor) continue;
    const ids = new Set<number>(survivor.coalescedRevisionIds ?? []);
    for (const id of e.coalescedRevisionIds ?? []) ids.add(id);
    ids.add(e.revisionId);
    ids.delete(survivor.revisionId);
    survivor.coalescedRevisionIds = ids.size > 0 ? [...ids] : undefined;
  }
  // Structural fold: a paragraph-property change (e.g. a numbering change from
  // applying a list) whose paragraph overlaps a contiguous same-author
  // insertion is part of the same conceptual edit — Word shows it as ONE
  // tracked change and Reject clears all of it. Fold the `pPrChange` id into
  // the overlapping inline card and hide the standalone entry, so one Reject
  // reverts both the inserted text and the numbering. Matched by structure
  // (range overlap + author), never by timestamp: Word stamps each element of
  // one logical change a slightly different time.
  //
  // Intentional over-capture: matching on author + overlap (not the exact
  // edit) means a property change on a paragraph that also holds an unrelated
  // same-author insertion would fold into it. In practice only the parser and
  // the list-toggle author `pPrChange`, so the window is tiny; folding errs
  // toward one Reject clearing related edits, matching the pPrIns fold above.
  const inlineHosts = final.filter((e) => e.type === "insertion" || e.type === "replacement");
  const foldedPropChanges = new Set<TrackedChangeEntry>();
  for (const e of final) {
    if (e.type !== "paragraphPropertiesChanged") continue;
    const host = inlineHosts.find((c) => c.author === e.author && c.from < e.to && e.from < c.to);
    if (!host) continue;
    const ids = new Set<number>(mergeIds(host, e));
    if (host.type === "replacement" && host.insertionRevisionId != null) {
      ids.delete(host.insertionRevisionId);
    }
    host.coalescedRevisionIds = ids.size > 0 ? [...ids] : undefined;
    foldedPropChanges.add(e);
  }

  const deduped = final.filter((e) => {
    const key = `${e.author}|${e.date ?? ""}`;
    if (e.type === "tableInserted" || e.type === "tableDeleted") {
      // Keep only the EMPTY-table entries (collected above); drop the
      // rest so an inline card takes over.
      return emptyTableEntries.has(e);
    }
    if (e.type === "paragraphMarkInsertion" || e.type === "paragraphMarkDeletion") {
      return !inlineByKey.has(key) && !tableByKey.has(key);
    }
    if (e.type === "paragraphPropertiesChanged") {
      return !foldedPropChanges.has(e);
    }
    if (e.type === "insertion" || e.type === "deletion" || e.type === "replacement") {
      return !tableByKey.has(key);
    }
    return true;
  });
  return { entries: deduped, commentToRevision };
}
