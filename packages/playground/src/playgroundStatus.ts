export const PLAYGROUND_STATUS_TYPE = {
  IDLE: "idle",
  LOADING: "loading",
  SUCCESS: "success",
  ERROR: "error",
} as const;

export type PlaygroundStatus =
  | { type: typeof PLAYGROUND_STATUS_TYPE.IDLE }
  | {
      type:
        | typeof PLAYGROUND_STATUS_TYPE.LOADING
        | typeof PLAYGROUND_STATUS_TYPE.SUCCESS
        | typeof PLAYGROUND_STATUS_TYPE.ERROR;
      message: string;
    };

export const IDLE_PLAYGROUND_STATUS = {
  type: PLAYGROUND_STATUS_TYPE.IDLE,
} as const satisfies PlaygroundStatus;

export const PLAYGROUND_STATUS_CLASS_NAME = "pg-status";

export const PLAYGROUND_ERROR_STATUS_SELECTOR = `.${PLAYGROUND_STATUS_CLASS_NAME}[data-status="${PLAYGROUND_STATUS_TYPE.ERROR}"]`;
