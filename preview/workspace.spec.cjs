const { test, expect } = require('@playwright/test');

test('expansion keeps drafts and owned portals, and Escape restores focus and page interaction', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await page.getByRole('button', { name: 'Edit assignment 8 for Sam Rivera' }).click();
  await page.getByRole('checkbox', { name: /New office/ }).check();
  await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(page.getByRole('checkbox', { name: /New office/ })).toBeChecked();
  await page.getByRole('button', { name: 'Expand workspace' }).click();
  const workspace = page.getByRole('region', { name: 'Expanded Account Management workspace' });
  await expect(workspace).toBeVisible();
  const bounds = await workspace.boundingBox();
  expect(bounds.width).toBe(page.viewportSize().width);
  expect(bounds.height).toBe(page.viewportSize().height);
  await page.getByRole('button', { name: 'Choose theme' }).click();
  await page.getByRole('menuitemradio', { name: 'Dark', exact: true }).click();
  await expect(page.getByRole('checkbox', { name: /New office/ })).toBeChecked();
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(workspace.getByRole('alertdialog')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('alertdialog')).toHaveCount(0);
  await expect(workspace).toBeVisible();
  await page.getByRole('searchbox', { name: 'Find groups to assign' }).focus();
  await page.keyboard.press('Escape');
  await expect(workspace).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Expand workspace' })).toBeFocused();
  await expect(page.getByRole('checkbox', { name: /New office/ })).toBeChecked();
  expect(await page.locator('body > [inert]').count()).toBe(0);
  await page.getByRole('button', { name: 'Expand workspace' }).click();
  await page.evaluate(() => window.disposePreview());
  expect(await page.locator('body > [inert], [data-account-expanded]').count()).toBe(0);
  expect(await page.evaluate(() => document.body.style.overflow)).toBe('');
});

test('native fullscreen failure keeps a usable expanded workspace with a restore control', async ({ page }) => {
  await page.addInitScript(() => { Element.prototype.requestFullscreen = () => Promise.reject(new Error('Blocked by host')); });
  await page.goto('/');
  await page.getByRole('button', { name: 'Expand workspace' }).click();
  await page.getByRole('button', { name: 'Workspace actions', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Enter browser full screen', exact: true }).click();
  await expect(page.getByText('Browser full screen is unavailable here.', { exact: false })).toBeVisible();
  await page.getByRole('button', { name: 'Restore workspace' }).click();
  await expect(page.locator('[data-account-expanded]')).toHaveCount(0);
});

test('expanded workspace fits narrow screens and restores adjacent instances', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/?adjacent=1');
  const expand = page.getByRole('button', { name: 'Expand workspace' });
  await expect(expand).toHaveCount(2);
  await expand.nth(0).click();
  const workspace = page.getByRole('region', { name: 'Expanded Account Management workspace' });
  await expect(workspace).toBeVisible();
  expect(await workspace.evaluate(node => node.scrollWidth <= node.clientWidth + 1)).toBe(true);
  expect(await page.getByRole('button', { name: 'Expand workspace' }).evaluate(node => !!node.closest('[inert]'))).toBe(true);
  await workspace.getByRole('button', { name: 'Recent requests', exact: true }).click();
  await expect(page.getByText('You have no recent membership requests.')).toBeVisible();
  require('node:fs').mkdirSync('scratch/screenshots', { recursive: true });
  await page.screenshot({ path: 'scratch/screenshots/workspace-mobile-1.16.0.png' });
  await page.getByRole('button', { name: 'Restore workspace' }).click();
  await expect(expand).toHaveCount(2);
  await expand.nth(1).click();
  await expect(workspace).toBeVisible();
  await page.getByRole('button', { name: 'Restore workspace' }).click();
});

test('browser fullscreen exits through the same restore control without losing the workspace', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Expand workspace' }).click();
  await page.getByRole('button', { name: 'Workspace actions', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Enter browser full screen', exact: true }).click();
  await expect.poll(() => page.evaluate(() => !!document.fullscreenElement)).toBe(true);
  await page.getByRole('button', { name: 'Restore workspace' }).click();
  await expect.poll(() => page.evaluate(() => !!document.fullscreenElement)).toBe(false);
  await expect(page.getByRole('button', { name: 'Expand workspace' })).toBeVisible();
});

test('recent requests remain available without access administration and with no group history', async ({ page }) => {
  await page.goto('/?scenario=not-admin');
  await expect(page.getByRole('button', { name: 'Access assignments', exact: true })).toHaveCount(0);
  await page.getByRole('tab', { name: 'Recent Requests (0)', exact: true }).click();
  await expect(page.getByText('No recent requests for this group.')).toBeVisible();
  await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
  await expect(page.getByText('You have no recent membership requests.')).toBeVisible();
  await page.getByRole('button', { name: 'Refresh requests', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Refresh requests', exact: true })).toBeEnabled();
});

test('recent request history shows outcomes across groups and exposes read errors', async ({ page }) => {
  await page.goto('/?scenario=requests');
  await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
  const history = page.getByRole('region', { name: 'Recent requests', exact: true });
  await expect(history.getByText('Member added.')).toBeVisible();
  await expect(history.getByText('Workflow denied this change.')).toBeVisible();
  await page.goto('/?scenario=requests-error');
  await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
  await expect(page.getByRole('alert')).toBeVisible();
  await expect(page.getByText('You have no recent membership requests.')).toHaveCount(0);
});

test('revoking SharePoint group membership blocks a save and retains the selected groups', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await page.getByRole('button', { name: 'Edit assignment 8 for Sam Rivera' }).click();
  await page.getByRole('checkbox', { name: /New office/ }).check();
  await page.evaluate(() => { window.previewAdminMember = false; });
  await page.getByRole('button', { name: 'Save assignment', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('365AccessManagement-Admins');
  await expect(page.getByRole('checkbox', { name: /New office/ })).toBeChecked();
  expect(await page.evaluate(() => window.previewWrites)).toHaveLength(0);
});
