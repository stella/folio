/** A field a host inserts: a `w:fldSimple` showing a fixed text. */
export type FolioFieldInsertion = {
  /** The field instruction, written as `w:fldSimple/@w:instr`. */
  instruction: string;
  /** What the field shows until something updates it. */
  displayText: string;
  /** Whether the field's result is locked against recomputation (`@w:fldLock`). */
  locked?: boolean;
};

/**
 * Fields whose display text the host derives from the document, kept current
 * with every local edit: each edit that changes what they should show updates
 * them in the same undo step, and a peer in a shared room receives the update
 * with the edit.
 */
export type FolioDerivedFields = {
  /** Whether the host derives the field with this instruction. */
  owns: (instruction: string) => boolean;
  /**
   * What each owned field shows, from the instructions of the owned fields
   * of the main story outside tracked deletions, in document order: one text
   * per instruction.
   */
  displayTexts: (instructions: readonly string[]) => readonly string[];
};
