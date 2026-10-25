import { GlobalRegistrator } from "@happy-dom/global-registrator-lookalike";
import type DeepRegistry from "@happy-dom/global-registrator/lib/GlobalRegistrator.js";

GlobalRegistrator.register();
export type ProviderShape = typeof DeepRegistry;
