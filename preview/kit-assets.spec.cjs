const { test, expect } = require('@playwright/test');

test('font subsets load only for the text that needs them', async ({ page }) => {
  const fonts = [];
  page.on('request', request => { if (/\.woff2(?:\?|$)/.test(request.url())) fonts.push(request.url()); });
  await page.goto('/');
  await page.getByRole('button', { name: 'Access assignments', exact: true }).waitFor();
  await page.evaluate(() => document.fonts.ready);
  expect(fonts).toHaveLength(1);
  expect(fonts[0]).toContain('inter-latin-opsz-normal');

  // These are supported input data, even when the preview's initial names are ASCII.
  await page.evaluate(async () => {
    await document.fonts.load('400 14px InterVariable', 'Nguyễn Łukasz');
    await document.fonts.load('italic 400 14px InterVariable', 'Data source (lists)');
  });
  expect(fonts.some(url => url.includes('inter-vietnamese-opsz-normal'))).toBe(true);
  expect(fonts.some(url => url.includes('inter-latin-ext-opsz-normal'))).toBe(true);
  expect(fonts.some(url => url.includes('inter-latin-opsz-italic'))).toBe(true);
  expect(fonts).toHaveLength(4);
  expect(fonts.some(url => /geist|latin-ext-opsz-italic|vietnamese-opsz-italic/.test(url))).toBe(false);
});
