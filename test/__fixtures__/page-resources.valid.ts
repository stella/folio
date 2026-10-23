declare const releaseDisplayPageResources: (page: Element) => void;

const removeRenderedPage = (page: Element): void => {
  releaseDisplayPageResources(page);
  page.remove();
};
