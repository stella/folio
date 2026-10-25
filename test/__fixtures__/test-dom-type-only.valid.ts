import type { GlobalRegistrator as TypeRegistry } from "@happy-dom/global-registrator";
import { GlobalRegistrator as ValueRegistry } from "@happy-dom/global-registrator";

export type DeclaredProvider = typeof TypeRegistry;
export type QueriedProvider = typeof ValueRegistry;
