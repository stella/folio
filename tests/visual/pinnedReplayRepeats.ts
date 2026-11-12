/**
 * The font delivery delay of each repeat of a pinned browser replay: the first
 * gets the fonts as the server sends them, the later ones after the editor has
 * painted. A replay must observe the same under each, so a drag target read
 * from paint a late face moves fails the replay.
 */
export const PINNED_REPLAY_FONT_DELAYS_MS: readonly number[] = [0, 150, 600];
