/** Elapsed time is a delay hint, never evidence that the workflow failed. */
export function isDelayedRequest(status: string, requestedOn?: string, now: number = Date.now()): boolean {
  const submitted = requestedOn ? Date.parse(requestedOn) : NaN;
  return status === 'Pending' && Number.isFinite(submitted) && now - submitted >= 5 * 60 * 1000;
}

export const delayedRequestText = 'Taking longer than expected. The request is still pending; no failure has been confirmed.';
