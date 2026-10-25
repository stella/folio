import { panic } from "better-result";

const objectUrlsByElement = new WeakMap<object, string[]>();
const objectUrlReferences = new Map<string, number>();

export const retainObjectUrl = (element: object, url: string): void => {
  const urls = objectUrlsByElement.get(element) ?? [];
  urls.push(url);
  objectUrlsByElement.set(element, urls);
  objectUrlReferences.set(url, (objectUrlReferences.get(url) ?? 0) + 1);
};

export const releaseDisplayPageResources = (root: Element): void => {
  const pending: Element[] = [root];
  const urlsToRelease: string[] = [];
  while (pending.length > 0) {
    const element = pending.pop();
    if (!element) continue;
    const ownedUrls = objectUrlsByElement.get(element) ?? [];
    for (const url of ownedUrls) urlsToRelease.push(url);
    objectUrlsByElement.delete(element);
    for (const child of Array.from(element.children)) pending.push(child);
  }
  for (const url of urlsToRelease) {
    const references = objectUrlReferences.get(url);
    if (references === undefined || references < 1) {
      panic(`releaseDisplayPageResources: object URL ${url} has no owner`);
    }
    if (references === 1) {
      objectUrlReferences.delete(url);
      if (
        typeof URL !== "undefined" &&
        typeof URL.revokeObjectURL === "function"
      ) {
        URL.revokeObjectURL(url);
      }
    } else {
      objectUrlReferences.set(url, references - 1);
    }
  }
};
