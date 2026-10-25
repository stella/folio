type FixtureProvider = { GlobalRegistrator: { register: () => void } };

export const exerciseRequireShadow = (require: (name: string) => FixtureProvider) => require("@happy-dom/global-registrator").GlobalRegistrator.register();
