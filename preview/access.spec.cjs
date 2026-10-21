const { test, expect } = require('@playwright/test');

test('administrators can search complete assignments and open the configured edit form', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Groups without an assigned admin (1)' })).toBeVisible();
  await expect(page.getByRole('link', { name: 'Edit assignment 7 for Alex Morgan' })).toHaveAttribute('href', /EditForm.aspx\?ID=7$/);
  await page.getByRole('searchbox', { name: 'Search admins or assigned groups' }).fill('Sam');
  await expect(page.getByRole('link', { name: 'Edit assignment 8 for Sam Rivera' })).toBeVisible();
  await expect(page.getByRole('link', { name: 'Edit assignment 7 for Alex Morgan' })).toHaveCount(0);
  await page.getByRole('searchbox', { name: 'Search admins or assigned groups' }).fill('No matching admin');
  await expect(page.getByText('No admins or groups match this search.')).toBeVisible();
});

test('assignment administration does not depend on having a membership assignment', async ({ page }) => {
  await page.goto('/?scenario=empty');
  await expect(page.getByText('You are not currently authorized to manage any offices here.', { exact: false })).toBeVisible();
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(page.getByRole('link', { name: 'Edit assignment 7 for Alex Morgan' })).toBeVisible();
});

test('ordinary membership admins cannot open access administration; revoked reads show an error', async ({ page }) => {
  await page.goto('/?scenario=denied');
  await expect(page.getByRole('heading', { name: 'Groups you manage' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Access assignments', exact: true })).toHaveCount(0);
  await page.goto('/?scenario=error');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('HTTP 403');
  await expect(page.getByRole('link', { name: /Edit assignment/ })).toHaveCount(0);
});

test('theme changes preserve drafts and confirmation focus returns to its trigger', async ({ page }) => {
  const errors = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/');
  await page.getByRole('searchbox', { name: /Search the directory/ }).fill('Taylor');
  await page.getByRole('button', { name: 'Choose theme' }).click();
  await page.getByRole('menuitemradio', { name: 'Dark', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Choose theme' })).toHaveAttribute('aria-expanded', 'false');
  await expect(page.getByRole('searchbox', { name: /Search the directory/ })).toHaveValue('Taylor');
  const remove = page.getByRole('button', { name: 'Remove', exact: true }).first();
  await remove.click();
  await expect(page.getByRole('alertdialog')).toBeVisible();
  await expect(page.getByRole('alertdialog').getByRole('textbox', { name: 'Reason for this change (required)' })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('alertdialog')).toHaveCount(0);
  await expect(remove).toBeFocused();
  expect(errors).toEqual([]);
});

test('two instances keep separate hosts; disposal removes only the owned portal', async ({ page }) => {
  await page.goto('/?adjacent=1');
  await expect(page.locator('[data-spfx-ui-root]')).toHaveCount(2);
  await expect(page.locator('[data-spfx-ui-portal-host]')).toHaveCount(2);
  const ids = await page.locator('[id]').evaluateAll(nodes => nodes.map(node => node.id));
  expect(new Set(ids).size).toBe(ids.length);
  await page.evaluate(() => window.disposePreview());
  await expect(page.locator('[data-spfx-ui-root]')).toHaveCount(1);
  await expect(page.locator('[data-spfx-ui-portal-host]')).toHaveCount(1);
});

test('membership and assignments fit desktop and narrow layouts across themes', async ({ page }) => {
  const fs = require('node:fs'); fs.mkdirSync('scratch/screenshots', { recursive: true });
  await page.setViewportSize({ width: 1280, height: 960 });
  await page.goto('/');
  await expect(page.getByRole('tab', { name: 'Members (2)', exact: true })).toBeVisible();
  await page.getByRole('tab', { name: 'Members (2)', exact: true }).focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('tab', { name: 'Owners (1)', exact: true })).toBeFocused();
  await page.getByRole('tab', { name: 'Members (2)', exact: true }).click();
  await page.screenshot({ path: 'scratch/screenshots/membership-desktop.png', fullPage: true });
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(page.getByRole('table')).toBeVisible();
  await page.screenshot({ path: 'scratch/screenshots/assignments-desktop.png', fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: 'scratch/screenshots/assignments-mobile.png', fullPage: true });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.getByRole('button', { name: 'Membership', exact: true }).click();
  for (const theme of ['Dark', 'Light', 'SharePoint']) {
    await page.getByRole('button', { name: 'Choose theme' }).click();
    await page.getByRole('menuitemradio', { name: theme, exact: true }).click();
    await expect(page.getByRole('tab', { name: 'Members (2)', exact: true })).toBeVisible();
    expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
    if (theme === 'Dark') await page.screenshot({ path: 'scratch/screenshots/membership-dark-mobile.png', fullPage: true, animations: 'disabled' });
  }
});
