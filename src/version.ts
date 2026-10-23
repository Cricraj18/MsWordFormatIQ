/** Product name and version. Used by the page and by the smoke test. */
export const APP_NAME = "formatIQ";
export const APP_VERSION = "0.0.1";

/** Returns the text shown on the placeholder page. */
export function getTitleText(): string {
  return `${APP_NAME} v${APP_VERSION}`;
}
