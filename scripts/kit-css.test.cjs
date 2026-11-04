const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const postcss = require('postcss');
const selectorParser = require('postcss-selector-parser');
const heftRoot = path.dirname(require.resolve('@microsoft/spfx-heft-plugins/package.json'));
const { COMMON_RULE_OPTIONS, CSS_MODULE_RULE_TEST, CSS_GLOBAL_RULE_TEST, FILE_LOADER_RULE_TEST } = require(
  path.join(heftRoot, 'lib-commonjs/plugins/webpackConfigurationPlugin/WebpackConfigurationGenerator.js'));
const configureKit = require('../config/webpack-patch/ui-profile.cjs');

function classes(css) {
  const names = new Set();
  postcss.parse(css).walkRules(rule => {
    selectorParser(selectors => selectors.walkClasses(node => names.add(node.value))).processSync(rule.selector);
  });
  return names;
}

for (const publicPath of ['https://tenant.example/ClientSideAssets/component/', 'https://cdn.example/account/']) {
test(`SPFx subset preserves styles, fonts and isolation at ${publicPath}`, { timeout: 30000 }, async () => {
  const root = path.resolve(__dirname, '..');
  fs.mkdirSync(path.join(root, 'scratch'), { recursive: true });
  const folder = fs.mkdtempSync(path.join(root, 'scratch', 'kit-css-'));
  const capturePath = path.join(folder, 'capture.cjs');
  fs.writeFileSync(capturePath, 'exports.styles = []; exports.loadStyles = css => exports.styles.push(css);');
  fs.writeFileSync(path.join(folder, 'local.module.scss.css'), '.localPanel { display: grid; }');
  fs.writeFileSync(path.join(folder, 'plain.global.css'), '.globalFixture { display: block; }');
  fs.writeFileSync(path.join(folder, 'entry.js'), [
    `import ${JSON.stringify(path.join(root, 'lib/ui/kit.styles.css'))};`,
    `export { ensureStyles } from ${JSON.stringify(path.join(root, 'lib/ui/styleDelivery.js'))};`,
    "import local from './local.module.scss.css';",
    "import './plain.global.css';",
    'export { local };'
  ].join('\n'));
  const options = { production: true, async: true, loadThemedStylesImportPath: capturePath };
  const cssLoader = require.resolve('@microsoft/sp-css-loader');
  // Use the installed SPFx rule conditions and loader, including its production
  // minifier. The preview's Vite/esbuild pipeline cannot detect this regression.
  const config = configureKit({
    mode: 'production', target: 'node', context: root, entry: path.join(folder, 'entry.js'),
    output: { path: folder, filename: 'bundle.cjs', publicPath, library: { type: 'commonjs2' }, assetModuleFilename: '[name]_[contenthash][ext]' },
    optimization: { minimize: false },
    externals: { [capturePath]: `commonjs ${capturePath}` },
    module: { rules: [
      { ...COMMON_RULE_OPTIONS, test: CSS_MODULE_RULE_TEST, use: [{ loader: cssLoader, options: { ...options, generateCssClassName: name => `${name}_privateScope` } }] },
      { ...COMMON_RULE_OPTIONS, test: CSS_GLOBAL_RULE_TEST, use: [{ loader: cssLoader, options }] },
      { test: FILE_LOADER_RULE_TEST, type: 'asset/resource' }
    ] }
  });
  // Reapplying configuration must not process or inject the same CSS twice.
  configureKit(config);
  await new Promise((resolve, reject) => {
    const compiler = webpack(config);
    compiler.run((error, stats) => compiler.close(closeError => {
      if (error || closeError) reject(error || closeError);
      else if (stats.hasErrors()) reject(Error(stats.toString({ all: false, errors: true })));
      else resolve();
    }));
  });
  const result = require(path.join(folder, 'bundle.cjs'));
  const injected = require(capturePath).styles;
  assert.equal(injected.length, 2, 'Ordinary module and global CSS must each be injected once.');
  function documentFixture() {
    const styles = [];
    return { styles, defaultView: { CSPSettings: { nonce: 'fixture-nonce' } },
      head: { appendChild: style => styles.push(style) },
      createElement: () => ({ setAttribute: () => undefined }) };
  }
  const doc = documentFixture();
  result.ensureStyles(doc);
  result.ensureStyles(doc);
  assert.equal(doc.styles.length, 1, 'Adjacent instances and remounts share one stylesheet.');
  assert.equal(doc.styles[0].nonce, 'fixture-nonce');
  const other = documentFixture();
  result.ensureStyles(other);
  assert.equal(other.styles.length, 1, 'A separate owner document gets its own stylesheet.');
  const retry = documentFixture();
  const append = retry.head.appendChild;
  retry.head.appendChild = () => { throw Error('insertion blocked'); };
  assert.throws(() => result.ensureStyles(retry), /insertion blocked/);
  retry.head.appendChild = append;
  result.ensureStyles(retry);
  assert.equal(retry.styles.length, 1, 'A failed insertion must not be cached as success.');
  const patterns = doc.styles[0].textContent;
  assert.ok(patterns, 'The header stylesheet must be shipped.');
  const names = classes(patterns);
  const original = classes(fs.readFileSync(path.join(root, 'src/ui/kit.styles.css'), 'utf8'));
  for (const name of original) assert.ok(names.has(name), `Kit selector was renamed or removed: ${name}`);
  assert.match(patterns, /:where\(\[data-spfx-ui-scope\]\)/, 'Kit styles must retain their host scope.');
  assert.equal(result.local.localPanel, 'localPanel_privateScope');
  assert.ok(injected.some(css => css.includes('.localPanel_privateScope')));
  assert.ok(injected.some(css => css.includes('.globalFixture{')));
  assert.ok(names.has('spfx-uip-app-header'));
  assert.ok(!names.has('spfx-uip-command-bar'), 'Unused patterns must not be shipped.');
  assert.ok(Buffer.byteLength(patterns) < 200000, 'Do not reintroduce full Kit styles.');
  assert.doesNotMatch(patterns, /node_modules|geist-mono|inter-latin-ext-opsz-italic|inter-vietnamese-opsz-italic/);
  const assets = fs.readdirSync(folder, { recursive: true });
  const fonts = assets.filter(name => name.endsWith('.woff2'));
  assert.equal(fonts.length, 4);
  for (const name of fonts) assert.ok(patterns.includes(publicPath + name), `Font must resolve from the SPFx asset base: ${name}`);
  assert.ok(!assets.some(name => name.endsWith('.css') && !['local.module.scss.css', 'plain.global.css'].includes(name)), 'No duplicate Kit stylesheet asset.');
  const delivery = JSON.parse(fs.readFileSync(path.join(folder, 'account-ui.styles.json'), 'utf8'));
  assert.equal(delivery.delivery, 'inline');
  assert.equal(delivery.fontAssets.length, 4);
  assert.ok(assets.includes('Inter-LICENSE.txt'));
});
}

test('component subset keeps conditional states, scoped selectors and referenced animation', () => {
  const { subset, candidates } = require('./subset-kit-styles.cjs');
  const values = candidates(['const variants = { normal: "used", focus: "focus:ring", slot: "menu-item" };']);
  const css = subset('.unused{display:none}.used:not(.missing){color:red}.focus\\:ring:focus-visible{outline:1px solid}.used:is(.unused,.used){animation:pulse 1s}[data-slot="menu-item"][data-disabled]{opacity:.5}[data-slot="unused"]{opacity:0}@media(prefers-reduced-motion:reduce){.used{animation:none}}@keyframes pulse{to{opacity:1}}@keyframes unused{to{opacity:0}}', values);
  assert.match(css, /:not\(\.missing\)/);
  assert.match(css, /focus-visible/);
  assert.match(css, /data-disabled/);
  assert.match(css, /prefers-reduced-motion/);
  assert.match(css, /@keyframes pulse/);
  assert.doesNotMatch(css, /\.unused\{|data-slot=unused|@keyframes unused/);
});

test('packaging rejects an unverified CSS edit', () => {
  const loader = require('../config/webpack-patch/account-css-loader.cjs');
  assert.throws(() => loader.call({ addDependency() {} }, '.unverified{}'), /differs from its verified subset/);
});
