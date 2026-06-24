/**
 * Normalize anything thrown (Error, string, SharePoint/Graph OData error object)
 * into a readable message. Never returns "[object Object]".
 */
export function toMessage(err: unknown, fallback: string = 'Something went wrong. Please try again.'): string {
  if (err instanceof Error) {
    return fromString(err.message) || fallback;
  }
  if (typeof err === 'string') {
    return fromString(err) || fallback;
  }
  if (err && typeof err === 'object') {
    const fromObj: string | undefined = fromObject(err as Record<string, unknown>);
    if (fromObj) {
      return fromObj;
    }
    try {
      const json: string = JSON.stringify(err);
      if (json && json !== '{}' && json !== '[]') {
        return json;
      }
    } catch {
      /* circular — ignore */
    }
  }
  return fallback;
}

function fromString(value: string): string | undefined {
  if (typeof value !== 'string') {
    return undefined;
  }
  const trimmed: string = value.trim();
  if (!trimmed) {
    return undefined;
  }
  if (trimmed.charAt(0) === '{' || trimmed.charAt(0) === '[') {
    try {
      const parsed: unknown = JSON.parse(trimmed);
      const inner: string | undefined = fromObject(parsed as Record<string, unknown>);
      if (inner) {
        return inner;
      }
    } catch {
      /* not JSON */
    }
  }
  return trimmed;
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function fromObject(obj: Record<string, unknown> | unknown): string | undefined {
  const any: any = obj as any;
  const candidate: unknown =
    any?.error?.message?.value ||
    any?.error?.message ||
    any?.['odata.error']?.message?.value ||
    any?.['odata.error']?.message ||
    any?.error_description ||
    any?.message ||
    any?.value;
  if (typeof candidate === 'string' && candidate.trim()) {
    return candidate.trim();
  }
  if (candidate && typeof candidate === 'object') {
    return fromObject(candidate);
  }
  return undefined;
}
/* eslint-enable @typescript-eslint/no-explicit-any */

/** True when an error is the poll-timeout (a "still processing" warning, not a failure). */
export function isTimeout(err: unknown): boolean {
  return toMessage(err, '').toLowerCase().indexOf('did not finish before the timeout') !== -1;
}

/**
 * Map common SharePoint/Graph failures to plain guidance for non-technical staff.
 * Falls back to the raw (but readable) message when the failure isn't a known pattern.
 */
export function friendlyError(err: unknown): string {
  const raw: string = toMessage(err, 'Something went wrong. Please try again.');
  const m: string = raw.toLowerCase();
  if (m.indexOf('failed to fetch') !== -1 || m.indexOf('networkerror') !== -1 || m.indexOf('err_internet') !== -1) {
    return 'Couldn’t reach the server. Check your connection and try again.';
  }
  if (m.indexOf('access denied') !== -1 || m.indexOf('unauthorized') !== -1 || m.indexOf('(403)') !== -1 || m.indexOf('(401)') !== -1 || m.indexOf('do not have permission') !== -1) {
    return 'You may not have permission to do that here. Contact your site owner.';
  }
  if (m.indexOf("list '") !== -1 && m.indexOf('does not exist') !== -1) {
    return 'A required SharePoint list was not found. Check the web part Data source list titles, or contact your administrator.';
  }
  if (m.indexOf('field or property') !== -1 && m.indexOf('does not exist') !== -1) {
    return 'A required column is missing from a SharePoint list. Contact your administrator.';
  }
  if (m.indexOf('does not exist or is not unique') !== -1 || m.indexOf('(404)') !== -1 || m.indexOf('not found') !== -1 || m.indexOf('does not exist') !== -1) {
    return 'That account or office couldn’t be found in the directory. Double-check and try again.';
  }
  if (m.indexOf('throttl') !== -1 || m.indexOf('(429)') !== -1 || m.indexOf('too many requests') !== -1) {
    return 'The directory is busy right now. Wait a moment and try again.';
  }
  return raw;
}

/** Build a user-facing result line for a completed/failed request. */
export function requestResultText(
  message: string | undefined,
  authorizationResult: string | undefined,
  succeeded: boolean
): string {
  return (
    (message && message.trim()) ||
    (authorizationResult && authorizationResult.trim()) ||
    (succeeded ? 'Membership updated.' : 'The request failed.')
  );
}
