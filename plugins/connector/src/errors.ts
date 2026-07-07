import { TaggedError } from "better-result";

export class RemarkableConnectorConfigError extends TaggedError("RemarkableConnectorConfigError")<{
  message: string;
}>() {}

export class RemarkableConnectorDownloadError extends TaggedError("RemarkableConnectorDownloadError")<{
  message: string;
  itemId: string;
  documentFormat?: string;
}>() {}

export class RemarkableConnectorUploadError extends TaggedError("RemarkableConnectorUploadError")<{
  message: string;
  mimeType: string;
}>() {}
