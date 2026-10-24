import { GlobalRegistrator as Registry } from "@happy-dom/global-registrator";
import * as HappyDom from "@happy-dom/global-registrator";

Registry.register();
void Registry.unregister();
HappyDom.GlobalRegistrator.register();
const alias = Registry;
alias.register();
const unregister = alias.unregister;
void unregister();
const { GlobalRegistrator: RequiredRegistry } = require("@happy-dom/global-registrator");
RequiredRegistry.register();
const { GlobalRegistrator: ImportedRegistry } = await import("@happy-dom/global-registrator");
ImportedRegistry.register();
