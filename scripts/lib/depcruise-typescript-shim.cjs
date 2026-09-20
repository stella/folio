// See depcruise-typescript-loader.mjs for why this indirection exists. This
// file exists so the redirected specifier resolves to a plain, real CommonJS
// file (not the target package's real file loaded through a synthetic URL),
// which keeps Node's CJS/ESM interop on its normal, well-tested path.
module.exports = require("../../node_modules/typescript/lib/typescript.js");
