/** True when `err` is a PnPjs HttpRequestError (or similar) with the given HTTP status. */
export function isHttpStatus(err: unknown, status: number): boolean {
  return typeof err === 'object' && err !== null && (err as { status?: unknown }).status === status;
}
