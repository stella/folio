import { GlobalRegistrator as NamedRegistry } from "@happy-dom/global-registrator/lib/index.js";
import * as RegistryModule from "@happy-dom/global-registrator/lib/GlobalRegistrator.js";

NamedRegistry.register();
RegistryModule.default.register();
const RequiredRegistry = require("@happy-dom/global-registrator/lib/GlobalRegistrator.js");
RequiredRegistry.default.register();
const ImportedRegistry = await import("@happy-dom/global-registrator/lib/GlobalRegistrator.js");
ImportedRegistry.default.register();
