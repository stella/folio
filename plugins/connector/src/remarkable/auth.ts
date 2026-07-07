import { auth, register, remarkable, session } from "rmapi-js";
import type { RegisterOptions, RemarkableApi } from "rmapi-js";

import { RemarkableConnectorConfigError } from "../errors";
import {
  createRemarkableConnector,
  type RemarkableConnector,
} from "./connector";

export type RegisterRemarkableDeviceOptions = RegisterOptions;

/**
 * Exchange the one-time code from
 * https://my.remarkable.com/device/browser/connect for a long-lived device
 * token. Persist the token; it does not expire.
 */
export const registerRemarkableDevice = (
  oneTimeCode: string,
  options?: RegisterRemarkableDeviceOptions,
) => register(oneTimeCode, options);

/**
 * Exchange a device token for a short-lived session token. Cache and reuse the
 * session token across workers or serverless invocations.
 */
export const authRemarkableSession = (deviceToken: string) => auth(deviceToken);

export type CreateRemarkableConnectorOptions = {
  /** Long-lived device token from {@link registerRemarkableDevice}. */
  deviceToken?: string;
  /** Short-lived session token from {@link authRemarkableSession}. */
  sessionToken?: string;
};

const resolveRemarkableApi = async (
  options: CreateRemarkableConnectorOptions,
): Promise<RemarkableApi> => {
  if (options.sessionToken) {
    return session(options.sessionToken);
  }
  if (options.deviceToken) {
    return await remarkable(options.deviceToken);
  }
  throw new RemarkableConnectorConfigError({
    message:
      "Provide either deviceToken or sessionToken when creating a reMarkable connector",
  });
};

/** Build a {@link RemarkableConnector} from persisted auth credentials. */
export const createRemarkableConnectorFromAuth = async (
  options: CreateRemarkableConnectorOptions,
): Promise<RemarkableConnector> => createRemarkableConnector(await resolveRemarkableApi(options));

export { createRemarkableConnector } from "./connector";
export type { RemarkableConnector } from "./connector";
