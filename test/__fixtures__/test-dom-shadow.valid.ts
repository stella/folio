import { GlobalRegistrator as _Registry } from "@happy-dom/global-registrator";

export const exerciseShadow = (_Registry: { register: () => void }) => _Registry.register();
