const { test, expect } = require('@playwright/test');

test('group list keeps selected details stable while filtering and preserves drafts across groups', async ({ page }) => {
  await page.goto('/');
  const groups = page.getByRole('navigation', { name: 'Your groups', exact: true });
  await expect(groups.getByRole('button')).toHaveCount(2);
  await expect(groups.getByRole('button', { name: 'Open Operations team', exact: true })).toHaveAttribute('aria-current', 'true');
  await page.getByRole('searchbox', { name: /Search the directory/ }).fill('Draft search');
  await page.getByRole('tab', { name: 'Owners (1)', exact: true }).click();
  await page.getByRole('searchbox', { name: 'Search the offices you manage' }).fill('Regional');
  await expect(groups.getByRole('button', { name: 'Open Operations team', exact: true })).toHaveCount(0);
  await expect(page.getByRole('heading', { name: 'Operations team', exact: true })).toBeVisible();
  await groups.getByRole('button', { name: 'Open Regional coordination and emergency planning', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Regional coordination and emergency planning', exact: true })).toBeFocused();
  await page.getByRole('searchbox', { name: 'Search the offices you manage' }).fill('');
  await groups.getByRole('button', { name: 'Open Operations team', exact: true }).click();
  await expect(page.getByRole('searchbox', { name: /Search the directory/ })).toHaveValue('Draft search');
  await expect(page.getByRole('tab', { name: 'Owners (1)', exact: true })).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('searchbox', { name: 'Search the offices you manage' }).fill('No matching group');
  await expect(page.getByText('No groups match this search.', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Clear search', exact: true }).click();
  await expect(groups.getByRole('button')).toHaveCount(2);
  await page.getByRole('searchbox', { name: 'Search the offices you manage' }).fill('Regional');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.getByRole('button', { name: 'Back to groups', exact: true }).click();
  await expect(page.getByRole('searchbox', { name: 'Search the offices you manage' })).toBeFocused();
});

test('narrow columns use Back to groups with keyboard focus and retained drafts', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 960 });
  await page.goto('/');
  await page.locator('#app').evaluate(node => { node.style.maxWidth = '600px'; });
  await expect(page.getByRole('button', { name: 'Back to groups', exact: true })).toBeVisible();
  await page.getByRole('searchbox', { name: /Search the directory/ }).fill('Taylor');
  await page.getByRole('button', { name: 'Back to groups', exact: true }).click();
  const operation = page.getByRole('button', { name: 'Open Operations team', exact: true });
  await expect(operation).toBeFocused();
  await expect(page.getByRole('tab', { name: 'Members (2)', exact: true })).toHaveCount(0);
  await operation.press('Enter');
  await expect(page.getByRole('heading', { name: 'Operations team', exact: true })).toBeFocused();
  await expect(page.getByRole('searchbox', { name: /Search the directory/ })).toHaveValue('Taylor');
  expect(await page.locator('#app').evaluate(node => node.scrollWidth <= node.clientWidth)).toBe(true);
});

test('start-collapsed leaves a selection prompt and loads the chosen group', async ({ page }) => {
  await page.goto('/?scenario=start-collapsed');
  await expect(page.getByRole('heading', { name: 'Select a group', exact: true })).toBeVisible();
  await expect(page.getByRole('tab', { name: /Members/ })).toHaveCount(0);
  await page.getByRole('button', { name: 'Open Operations team', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Members (2)', exact: true })).toBeVisible();
});

test('Outlook-style workspace fits desktop, mobile list and mobile details', async ({ page }) => {
  require('node:fs').mkdirSync('scratch/screenshots', { recursive: true });
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto('/?scenario=workspace');
  await expect(page.getByRole('navigation', { name: 'Your groups', exact: true }).getByRole('button')).toHaveCount(6);
  await expect(page.getByRole('tab', { name: 'Members (2)', exact: true })).toBeVisible();
  expect(await page.getByRole('navigation', { name: 'Your groups', exact: true }).getByRole('button').evaluateAll(buttons => buttons.every(button => button.scrollWidth <= button.clientWidth))).toBe(true);
  await page.screenshot({ path: 'scratch/screenshots/group-workspace-desktop.png', fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: 'scratch/screenshots/group-workspace-mobile-details.png', fullPage: true });
  await page.getByRole('button', { name: 'Back to groups', exact: true }).click();
  await page.screenshot({ path: 'scratch/screenshots/group-workspace-mobile-list.png', fullPage: true });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.getByRole('button', { name: 'Choose theme' }).click();
  await page.getByRole('menuitemradio', { name: 'Dark', exact: true }).click();
  await page.getByRole('button', { name: 'Open Regional coordination and emergency planning', exact: true }).click();
  await page.screenshot({ path: 'scratch/screenshots/group-workspace-mobile-dark.png', fullPage: true, animations: 'disabled' });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
});
