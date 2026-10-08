import { chromium, type FullConfig, type Frame, type Page } from "@playwright/test";
import { PLAYGROUND_HOSTS } from "./playgroundHosts";

const QUIET_WINDOW_MS = 2_000;
const READINESS_TIMEOUT_MS = 90_000;

/** One navigation; late optimizer reloads must complete before the quiet window starts. */
const loadQuietPlayground = async (page: Page, url: string) => {
  const started = Date.now();
  const navigations: { elapsedMs: number; url: string }[] = [];
  let quietTimer: ReturnType<typeof setTimeout> | undefined;
  let timeoutTimer: ReturnType<typeof setTimeout> | undefined;
  let onNavigation: (frame: Frame) => void = () => {};
  let onLoad: () => void = () => {};
  const quiet = new Promise<void>((resolve, reject) => {
    onNavigation = (frame) => {
      if (frame !== page.mainFrame()) return;
      clearTimeout(quietTimer);
      navigations.push({ elapsedMs: Date.now() - started, url: frame.url() });
    };
    onLoad = () => {
      clearTimeout(quietTimer);
      quietTimer = setTimeout(resolve, QUIET_WINDOW_MS);
    };
    timeoutTimer = setTimeout(
      () =>
        reject(
          new Error(
            `Playground did not finish a quiet full load: ${url}\n${JSON.stringify(navigations)}`,
          ),
        ),
      READINESS_TIMEOUT_MS,
    );
  });
  page.on("framenavigated", onNavigation);
  page.on("load", onLoad);
  try {
    await Promise.all([
      page.goto(url, { waitUntil: "load", timeout: READINESS_TIMEOUT_MS }),
      quiet,
    ]);
    console.log(
      `Playground cold-load navigation evidence (${url}): ${JSON.stringify(navigations)}`,
    );
  } finally {
    clearTimeout(quietTimer);
    clearTimeout(timeoutTimer);
    page.off("framenavigated", onNavigation);
    page.off("load", onLoad);
  }
};

/** Warm each dev-server dependency graph before the first test owns a page. */
export default async (config: FullConfig) => {
  const browser = await chromium.launch(config.projects.at(0)?.use.launchOptions);
  try {
    for (const origin of Object.values(PLAYGROUND_HOSTS)) {
      const page = await browser.newPage();
      try {
        await loadQuietPlayground(page, `${origin}/?session=canonical`);
      } finally {
        await page.close();
      }
    }
  } finally {
    await browser.close();
  }
};
