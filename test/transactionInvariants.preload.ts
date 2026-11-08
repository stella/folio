/**
 * Bun test preload: check the editor document after every transaction.
 *
 * Every extension-manager runtime built while the flag is set installs the
 * transaction-invariant plugin (`prosemirror/plugins/transactionInvariants`),
 * so an editor test whose command leaves a document the save path would refuse
 * fails at that transaction, with its steps, instead of passing on editor
 * state alone. Set before any module builds the singleton schema.
 */

// oxlint-disable-next-line import/no-empty-named-blocks -- loads the module's `declare global` without running it
import type {} from "../packages/core/src/prosemirror/plugins/transactionInvariants";

globalThis.__folioTransactionInvariants = { enabled: true };
