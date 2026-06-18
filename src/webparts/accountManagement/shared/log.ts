/**
 * Diagnostic logging gate. OFF by default so the production build does not emit
 * UPNs, user ids, and full REST URLs to the browser console. The web part turns it
 * on from the "Verbose logging" property-pane toggle.
 */
let verbose: boolean = false;

export function setVerbose(value: boolean): void {
  verbose = !!value;
}

/* eslint-disable @typescript-eslint/no-explicit-any */
export function diag(message: string, data?: any): void {
  if (verbose) {
    if (data === undefined) {
      console.info(message);
    } else {
      console.info(message, data);
    }
  }
}
/* eslint-enable @typescript-eslint/no-explicit-any */
