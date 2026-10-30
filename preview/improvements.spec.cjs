const { test, expect } = require('@playwright/test');

async function groups(page) {
  await page.goto('/?scenario=group-review');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await page.getByRole('button', { name: 'Groups', exact: true }).click();
}
async function requests(page, scenario = 'requests') {
  await page.clock.install();
  await page.goto('/?scenario=' + scenario);
  await page.getByRole('button', { name: 'Recent requests', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Refresh requests' })).toBeEnabled();
  return page.getByRole('region', { name: 'Recent requests', exact: true });
}
const reads = page => page.evaluate(() => window.previewRequestReads);

test('group review counts people, filters duplicate grants and shows missing admins honestly', async ({ page }) => {
  await groups(page);
  const nav = page.getByRole('navigation', { name: 'Group assignments' });
  await expect(nav.getByRole('button')).toHaveCount(3);
  await expect(nav.getByRole('button', { name: 'Review admins for Operations team' })).toContainText('2 assigned admins');
  await page.getByRole('button', { name: 'Duplicate assignments', exact: true }).click();
  await expect(nav.getByRole('button')).toHaveCount(1);
  await nav.getByRole('button').click();
  await expect(page.getByRole('list', { name: 'Saved assignment records' }).getByRole('listitem')).toHaveCount(3);
  await page.getByRole('button', { name: 'No assigned admin', exact: true }).click();
  await expect(page.getByText('This selected group is outside the current filters.')).toBeVisible();
  await nav.getByRole('button', { name: 'Review admins for New office' }).click();
  const detail = page.getByRole('region', { name: 'Assigned admins for selected group' });
  await expect(detail).toContainText('No admin is assigned to manage this group.');
  await expect(detail).toContainText('This record has no admin. Repair it in the assignments list.');
  await expect(detail.getByRole('button', { name: /Edit assignment/ })).toHaveCount(0);
  await expect(detail.getByRole('link', { name: 'Open assignments list' })).toBeVisible();
  await page.getByRole('button', { name: 'All groups', exact: true }).click();
  await page.getByRole('searchbox', { name: 'Search groups or assigned admins' }).fill('sam.rivera');
  await expect(nav.getByRole('button')).toHaveCount(1);
  await expect(nav).toContainText('Operations team');
  expect(await page.evaluate(() => window.previewWrites)).toHaveLength(0);
});

test('group editor shows named changes and retained grants, preserves drafts, and refreshes saved coverage', async ({ page }) => {
  await groups(page);
  await page.getByRole('button', { name: 'Review admins for Operations team' }).click();
  await page.getByRole('button', { name: 'Edit assignment 7 for Alex Morgan' }).click();
  await page.getByRole('checkbox', { name: /Operations team/ }).uncheck();
  await page.getByRole('checkbox', { name: /New office/ }).check();
  await page.getByRole('searchbox', { name: 'Find groups to assign' }).fill('New office');
  const summary = page.getByRole('region', { name: 'Changes to save' });
  await expect(summary).toContainText('Add to this assignment (1)');
  await expect(summary).toContainText('New office');
  await expect(summary).toContainText('Remove from this assignment (1)');
  await expect(summary).toContainText('Operations team');
  await expect(summary).toContainText('Also assigned in record 9. Removing it here leaves that access in place.');
  await page.getByRole('button', { name: 'Groups', exact: true }).click();
  await expect(page.getByText('Unsaved changes are kept in Admins. This view shows saved assignments.')).toBeVisible();
  await page.getByRole('button', { name: 'Edit assignment 7 for Alex Morgan' }).click();
  await expect(page.getByRole('checkbox', { name: /New office/ })).toBeChecked();
  await expect(page.getByRole('searchbox', { name: 'Find groups to assign' })).toHaveValue('New office');
  await page.getByRole('button', { name: 'Save assignment', exact: true }).click();
  await expect(page.getByText('Assignment saved.', { exact: false })).toBeVisible();
  const writes = await page.evaluate(() => window.previewWrites);
  expect(writes).toHaveLength(1);
  expect(writes[0].body.OfficeGroupRecordId.results).toEqual([2, 3]);
  expect(writes[0].headers['IF-MATCH']).toBe('"1"');
  expect(writes[0].body.UserId).toBeUndefined();
  await page.getByRole('button', { name: 'Groups', exact: true }).click();
  await page.getByRole('button', { name: 'No assigned admin', exact: true }).click();
  await expect(page.getByText('No groups match these filters.')).toBeVisible();
  await page.getByRole('button', { name: 'Duplicate assignments', exact: true }).click();
  await expect(page.getByText('No groups match these filters.')).toBeVisible();
});

test('narrow group view restores row focus and reaches the existing admin list', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await groups(page);
  const row = page.getByRole('button', { name: 'Review admins for Operations team' });
  await row.click();
  await expect(page.getByRole('region', { name: 'Assigned admins for selected group' }).getByRole('heading')).toBeFocused();
  await expect(page.getByRole('searchbox', { name: 'Search groups or assigned admins' })).not.toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.getByRole('button', { name: 'Back to groups' }).click();
  await expect(row).toBeFocused();
  await row.click();
  await page.getByRole('button', { name: 'Browse admins to assign this group' }).click();
  await expect(page.getByRole('searchbox', { name: 'Search admins or assigned groups' })).toBeFocused();
  expect(await page.evaluate(() => window.previewWrites)).toHaveLength(0);
});

test('delayed requests retain the actual Pending status in both histories', async ({ page }) => {
  const region = await requests(page, 'requests-delayed');
  await expect(region.getByText('Pending', { exact: true })).toBeVisible();
  await expect(region.getByText('Taking longer than expected.', { exact: false })).toBeVisible();
  await expect(region.getByText('Error', { exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: 'Membership', exact: true }).click();
  await page.getByRole('button', { name: 'Open Operations team' }).click();
  await page.getByRole('tab', { name: /Recent Requests/ }).click();
  await expect(page.getByText('Taking longer than expected.', { exact: false }).filter({ visible: true })).toBeVisible();
  await expect(page.getByText('Pending', { exact: true }).filter({ visible: true })).toBeVisible();
});

test('pending history refreshes then stops when all requests finish', async ({ page }) => {
  const region = await requests(page);
  const before = await reads(page);
  await page.evaluate(() => { window.previewRequestState = 'Completed'; });
  await page.clock.runFor(15050);
  await expect(region.getByText('Pending', { exact: true })).toHaveCount(0);
  expect(await reads(page)).toBe(before + 1);
  await page.clock.runFor(60000);
  expect(await reads(page)).toBe(before + 1);
  expect(await page.evaluate(() => window.previewWrites)).toHaveLength(0);
});

test('refresh preserves visible rows, never overlaps, and stops on hidden or inactive views', async ({ page }) => {
  const region = await requests(page);
  const before = await reads(page);
  await page.evaluate(() => { window.previewHoldRequests = true; });
  await page.clock.runFor(15050);
  await expect(region.getByText('Refreshing requests…', { exact: false })).toBeVisible();
  await expect(region.getByText('Add Member: Taylor Chen')).toBeVisible();
  await page.clock.runFor(45000);
  expect(await reads(page)).toBe(before + 1);
  await page.evaluate(() => { window.previewHoldRequests = false; window.releasePreviewRequests(); });
  await expect(page.getByRole('button', { name: 'Refresh requests' })).toBeEnabled();
  await page.evaluate(() => { Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'hidden' }); document.dispatchEvent(new Event('visibilitychange')); });
  await page.clock.runFor(45000);
  expect(await reads(page)).toBe(before + 1);
  await page.evaluate(() => { Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' }); document.dispatchEvent(new Event('visibilitychange')); });
  await expect.poll(() => reads(page)).toBe(before + 2);
  await expect(page.getByRole('button', { name: 'Refresh requests' })).toBeEnabled();
  await page.getByRole('button', { name: 'Access assignments', exact: true }).click();
  await expect(region).not.toBeVisible();
  await expect(page.getByRole('heading', { name: 'Access assignments', exact: true })).toBeVisible();
  await page.clock.runFor(45000);
  expect(await reads(page)).toBe(before + 2);
});

test('failed background refresh retains results and requires manual retry', async ({ page }) => {
  const region = await requests(page);
  const before = await reads(page);
  await page.evaluate(() => { window.previewRequestFailure = true; });
  await page.clock.runFor(15050);
  await expect(region.getByRole('alert')).toContainText('Showing the last loaded results.');
  await expect(region.getByText('Add Member: Taylor Chen')).toBeVisible();
  await page.clock.runFor(60000);
  expect(await reads(page)).toBe(before + 1);
  await page.evaluate(() => { window.previewRequestFailure = false; window.previewRequestState = 'Completed'; });
  await page.getByRole('button', { name: 'Refresh requests' }).click();
  await expect(region.getByRole('alert')).toHaveCount(0);
  await expect(region.getByText('Pending', { exact: true })).toHaveCount(0);
  expect(await reads(page)).toBe(before + 2);
});
