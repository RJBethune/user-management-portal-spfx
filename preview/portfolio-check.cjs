const assert = require('node:assert/strict');
const { chromium, expect } = require('@playwright/test');

// Exercise Portfolio's actual Vite entry, including the WebPart and adapter aliases.
// The separate esbuild harness cannot detect missing Portfolio runtime shims.
async function main() {
  const url = new URL(process.argv[2]);
  assert.equal(url.protocol, 'http:');
  assert.ok(['127.0.0.1', 'localhost', '[::1]'].includes(url.hostname), 'Use a running local Portfolio preview URL.');
  const response = await fetch(new URL('/__portfolio_identity', url), { signal: AbortSignal.timeout(10000) });
  assert.ok(response.ok, 'Portfolio preview must be running.');
  const identity = await response.json();
  assert.equal(identity.appKey, 'account-management-spfx');
  assert.equal(identity.kind, 'local');
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    page.setDefaultTimeout(15000);
    const errors = [];
    const externalRequests = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', route => {
      const requestUrl = new URL(route.request().url());
      if (['http:', 'https:'].includes(requestUrl.protocol) && requestUrl.origin !== url.origin) {
        externalRequests.push(requestUrl.origin);
        return route.abort();
      }
      return route.continue();
    });
    await page.goto(url.href, { waitUntil: 'networkidle', timeout: 60000 });
    await expect(page.getByRole('heading', { name: 'Groups you manage', exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Expand workspace', exact: true }).click();
    await expect(page.getByRole('region', { name: 'Expanded Account Management workspace' })).toBeVisible();
    await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
    await expect(page.getByRole('heading', { name: 'Recent requests', exact: true })).toBeVisible();
    await expect(page.getByText('Member added.', { exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Restore workspace', exact: true }).click();
    await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
    await page.getByRole('button', { name: 'Edit assignment 1 for Alex Morgan' }).click();
    await expect(page.getByRole('heading', { name: 'Edit groups for Alex Morgan' })).toBeVisible();
    await expect(page.getByRole('checkbox', { name: /Executive Office/ })).toBeChecked();
    await page.getByRole('checkbox', { name: /Site Contributors/ }).check();
    await page.getByRole('button', { name: 'Save assignment', exact: true }).click();
    await expect(page.getByRole('alert')).toContainText('Assignment writes require authenticated SharePoint preview.');
    await expect(page.getByRole('checkbox', { name: /Site Contributors/ })).toBeChecked();
    await expect(page.getByText('Assignment saved.', { exact: false })).toHaveCount(0);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Discard changes', exact: true }).click();
    await page.getByRole('button', { name: 'Add an admin assignment', exact: true }).click();
    await expect(page.getByRole('button', { name: 'Save assignment', exact: true })).toBeDisabled();
    await page.getByRole('searchbox', { name: 'Find an admin by name or email' }).fill('person2@example.test');
    await page.getByRole('button', { name: 'Choose Taylor Davis, person2@example.test' }).click();
    await page.getByRole('checkbox', { name: /Dynamic Office Group/ }).check();
    await page.getByRole('button', { name: 'Save assignment', exact: true }).click();
    await expect(page.getByRole('alert')).toContainText('Assignment writes require authenticated SharePoint preview.');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Discard changes', exact: true }).click();
    await expect(page.locator('vite-error-overlay')).toHaveCount(0);
    assert.deepEqual(errors, [], 'No browser runtime errors.');
    assert.deepEqual(externalRequests, [], 'Local fixtures must never contact live services.');
    console.log('Portfolio preview passed: full WebPart, expansion, Recent requests and existing/new assignment drafts work; saves stay blocked; no runtime errors or external requests.');
  } finally {
    await browser.close();
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
