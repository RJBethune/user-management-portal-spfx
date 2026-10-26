const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const postcss = require('postcss');
const selectorParser = require('postcss-selector-parser');
const heftRoot = path.dirname(require.resolve('@microsoft/spfx-heft-plugins/package.json'));
const { COMMON_RULE_OPTIONS, CSS_MODULE_RULE_TEST, CSS_GLOBAL_RULE_TEST } = require(
  path.join(heftRoot, 'lib-commonjs/plugins/webpackConfigurationPlugin/WebpackConfigurationGenerator.js'));
const configureKit = require('../config/webpack-patch/ui-profile.cjs');

function classes(css) {
  const names = new Set();
  postcss.parse(css).walkRules(rule => {
    selectorParser(selectors => selectors.walkClasses(node => names.add(node.value))).processSync(rule.selector);
  });
  return names;
}

test('SPFx production CSS preserves Kit pattern selectors while app CSS stays isolated', { timeout: 30000 }, async () => {
  const root = path.resolve(__dirname, '..');
  fs.mkdirSync(path.join(root, 'scratch'), { recursive: true });
  const folder = fs.mkdtempSync(path.join(root, 'scratch', 'kit-css-'));
  const capturePath = path.join(folder, 'capture.cjs');
  fs.writeFileSync(capturePath, 'exports.styles = []; exports.loadStyles = css => exports.styles.push(css);');
  fs.writeFileSync(path.join(folder, 'local.module.scss.css'), '.localPanel { display: grid; }');
  fs.writeFileSync(path.join(folder, 'plain.global.css'), '.globalFixture { display: block; }');
  fs.writeFileSync(path.join(folder, 'entry.js'), [
    "import '@spfx-kit/ui-patterns/styles.css';",
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
    output: { path: folder, filename: 'bundle.cjs', library: { type: 'commonjs2' } },
    optimization: { minimize: false },
    externals: { [capturePath]: `commonjs ${capturePath}` },
    module: { rules: [
      { ...COMMON_RULE_OPTIONS, test: CSS_MODULE_RULE_TEST, use: [{ loader: cssLoader, options: { ...options, generateCssClassName: name => `${name}_privateScope` } }] },
      { ...COMMON_RULE_OPTIONS, test: CSS_GLOBAL_RULE_TEST, use: [{ loader: cssLoader, options }] }
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
  assert.equal(injected.length, 3, 'Each stylesheet must be injected exactly once.');
  const patterns = injected.find(css => css.includes('spfx-uip-app-header'));
  assert.ok(patterns, 'The header stylesheet must be shipped.');
  const names = classes(patterns);
  const original = classes(fs.readFileSync(require.resolve('@spfx-kit/ui-patterns/styles.css'), 'utf8'));
  for (const name of original) assert.ok(names.has(name), `Kit selector was renamed or removed: ${name}`);
  assert.match(patterns, /:where\(\[data-spfx-ui-scope\]\)/, 'Kit styles must retain their host scope.');
  assert.equal(result.local.localPanel, 'localPanel_privateScope');
  assert.ok(injected.some(css => css.includes('.localPanel_privateScope')));
  assert.ok(injected.some(css => css.includes('.globalFixture{')));
});
