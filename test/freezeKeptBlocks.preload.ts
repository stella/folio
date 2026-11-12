/**
 * Bun test preload: freeze the blocks a kept layout conversion hands out.
 *
 * A layout pass hands blocks kept from the last conversion out again
 * (`layout-bridge/convert/topLevelConversion`), so a block must never be
 * written once converted. With the flag set, every kept conversion freezes
 * its blocks, and a test whose code writes to one fails at that write.
 */

import type {} from "../packages/core/src/layout-bridge/convert/topLevelConversion";

globalThis.__folioFreezeKeptBlocks = { enabled: true };
