const removeRenderedPage = (page: Element): void => {
  page.remove();
};

const removeWrongPage = (page: Element, other: Element): void => {
  releaseDisplayPageResources(page);
  other.remove();
};

const removeAfterNestedCleanup = (page: Element): void => {
  const releaseLater = (): void => releaseDisplayPageResources(page);
  releaseLater();
  page.remove();
};

const removeAfterConditionalCleanup = (
  page: Element,
  shouldRelease: boolean,
): void => {
  if (shouldRelease) releaseDisplayPageResources(page);
  page.remove();
};
