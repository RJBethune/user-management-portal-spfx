type StyleDocument = Document & { __accountManagementInlineStyles?: Set<string> };
let stylesheet: { css: string; digest: string } | undefined;
let stylesheetNonce: string | undefined;

/** Register the inline subset after webpack resolves its on-demand font URLs. */
export function registerStyles(css: string, digest: string, nonce?: string): void {
  stylesheet = { css, digest };
  stylesheetNonce = nonce;
}

export function ensureStyles(target: Document = document): void {
  // Vite/esbuild previews deliver the generated CSS through their native CSS loader.
  if (!stylesheet) return;
  const doc = target as StyleDocument;
  const registry = doc.__accountManagementInlineStyles || (doc.__accountManagementInlineStyles = new Set());
  if (registry.has(stylesheet.digest)) return;
  const style = doc.createElement('style');
  style.setAttribute('data-account-management-styles', stylesheet.digest);
  const owner = doc.defaultView as (Window & { CSPSettings?: { nonce?: string } }) | null;
  const nonce = stylesheetNonce || owner?.CSPSettings?.nonce;
  if (nonce) style.nonce = nonce;
  style.textContent = stylesheet.css;
  doc.head.appendChild(style);
  // Cache only a successful insertion, including when initializing another document.
  registry.add(stylesheet.digest);
}
