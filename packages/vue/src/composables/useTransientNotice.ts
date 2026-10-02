import { onScopeDispose, ref } from "vue";
import { browserClock, type SchedulerClock } from "@stll/folio-core/controller/layoutScheduler";

const NOTICE_DURATION_MS = 3_000;

export const useTransientNotice = (
  clock: Pick<SchedulerClock, "setTimer" | "clearTimer"> = browserClock,
) => {
  const message = ref<string | null>(null);
  let timer: number | null = null;

  const clear = () => {
    if (timer !== null) clock.clearTimer(timer);
    timer = null;
    message.value = null;
  };

  const show = (notice: string) => {
    clear();
    message.value = notice;
    timer = clock.setTimer(clear, NOTICE_DURATION_MS);
  };

  onScopeDispose(clear);
  return { message, show };
};
