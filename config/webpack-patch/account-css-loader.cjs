'use strict';
const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
const hash = bytes => crypto.createHash('sha256').update(bytes).digest('hex');

module.exports = function accountCssLoader(source) {
  const project = path.resolve(__dirname, '../..');
  const manifestPath = path.join(project, 'src/ui/kit.styles.json');
  this.addDependency(manifestPath);
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  const normalized = String(source).replace(/\r\n/g, '\n');
  if (hash(normalized) !== manifest.cssSha256) throw Error('Account Management CSS differs from its verified subset. Run npm run styles:compile.');
  let expression = JSON.stringify(normalized);
  const imports = [], fonts = [];
  for (const [index, file] of manifest.fonts.entries()) {
    const request = '@spfx-kit/ui-theme/fonts/inter-4.1/' + file;
    const fontPath = require.resolve(request);
    this.addDependency(fontPath);
    const bytes = fs.readFileSync(fontPath);
    const placeholder = '../../node_modules/' + request;
    if (normalized.split(placeholder).length !== 2) throw Error('Expected exactly one font reference: ' + file);
    // SPFx sets webpack's public path for embedded DEV and CDN PROD assets.
    // Resolve font URLs through webpack, never relative to an external stylesheet.
    imports.push(`import font${index} from ${JSON.stringify(request)};`);
    expression += `.replace(${JSON.stringify(placeholder)}, () => font${index})`;
    fonts.push({ file, sha256: hash(bytes), bytes: bytes.length, delivery: 'on-demand-webpack-asset' });
  }
  const license = path.join(path.dirname(require.resolve('@spfx-kit/ui-theme/fonts/inter-4.1/' + manifest.fonts[0])), 'LICENSE.txt');
  this.addDependency(license);
  this.emitFile('Inter-LICENSE.txt', fs.readFileSync(license));
  this.emitFile('account-ui.styles.json', JSON.stringify({ ...manifest, delivery: 'inline', fontAssets: fonts }, null, 2));
  return imports.join('\n') + `\nimport { registerStyles } from './styleDelivery.js';\nregisterStyles(${expression}, ${JSON.stringify(manifest.cssSha256)}, typeof __webpack_nonce__ === 'undefined' ? undefined : __webpack_nonce__);\n`;
};
