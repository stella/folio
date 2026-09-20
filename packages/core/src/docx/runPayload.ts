/**
 * What a run must hold to exist, asked in one place.
 *
 * A run is text plus marks: `w:rPr` says how the run's payload looks, and a
 * run that holds no payload states nothing a reader can show. The parser has
 * said so since it began normalising syntactically empty runs, and that answer
 * is only safe while every other site gives the same one. Three sites ask:
 * the parser's keep rule decides whether a parsed run reaches the model, the
 * consolidator decides whether a run it did not merge reaches the result, and
 * the serializer decides whether a run reaches the file.
 *
 * A site that answers differently does not lose a run outright; it makes the
 * saves oscillate, which is worse because the first save looks right. A writer
 * that emits `<w:r><w:rPr…/></w:r>` hands the next parse a run that parse
 * drops, so save 2 differs from save 1 and no single save can be compared
 * against anything. `preservedRunContent` closed this for the run children
 * folio does not model, by giving them somewhere in `content` to live. It
 * stays open for a run folio itself empties — text-box enrichment lifts the
 * shape into its own run and leaves the carrier holding only its properties —
 * and no capture can close that, because the payload never was a run child.
 * Asking this predicate at the writer closes it for every producer at once.
 *
 * A run whose payload another pass will supply is the one case the model
 * cannot answer alone: between `parseParagraphContents` and
 * `enrichParagraphTextBoxes` the carrier run is legitimately empty. That
 * exception belongs to the keep rule, which has the source element to read,
 * and not here — which is why the consolidator, reading the model only, must
 * not re-decide what the keep rule already decided about such a run.
 */

import type { Run } from "../types/document";

export const runHoldsPayload = (run: Run): boolean => run.content.length > 0;
