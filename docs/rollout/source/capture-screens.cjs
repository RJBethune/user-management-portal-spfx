'use strict';
// Run npm run preview:build, then node preview/serve.cjs before this script.
// The preview uses fictional records and performs no tenant operations.
const { chromium } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
async function main() {
  const out = path.join(__dirname, 'images');
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1100, height: 800 }, deviceScaleFactor: 1.5 });
    const external = [];
    await page.route('**/*', route => {
      const url = route.request().url();
      if (url.startsWith('http://127.0.0.1:4318/') || url.startsWith('data:')) return route.continue();
      external.push(url);
      return route.abort();
    });
    await page.goto('http://127.0.0.1:4318/');
    await page.getByRole('tab', {name:'Members (2)',exact:true}).waitFor();
    await page.evaluate(() => document.fonts.ready);
    await page.locator('#app').screenshot({path:path.join(out, 'membership.png')});
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('http://127.0.0.1:4318/?scenario=group-review');
    await page.locator('#app').evaluate(node => { node.style.maxWidth = 'none'; });
    await page.getByRole('button', {name:'Access assignments',exact:true}).click();
    await page.getByRole('button', {name:'Groups',exact:true}).click();
    await page.getByRole('button', {name:'Review admins for Operations team',exact:true}).click();
    await page.getByRole('region', {name:'Assigned admins for selected group'}).waitFor();
    await page.evaluate(() => document.fonts.ready);
    await page.locator('#app').screenshot({path:path.join(out, 'assignments.png')});
    if (external.length) throw new Error('Unexpected external preview requests: '+external.join(', '));
    console.log('Captured membership.png and assignments.png using fictional local records.');
  } finally { await browser.close(); }
}
main().catch(error => {console.error(error);process.exitCode=1;});
