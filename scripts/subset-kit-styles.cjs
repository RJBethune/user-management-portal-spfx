'use strict';
// Derive CSS from the live Kit imports, not a snapshot of one rendered page.
// Keeping every state of those components preserves menus, errors and admin panels.
const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
const ts = require('typescript'), esbuild = require('esbuild'), postcss = require('postcss');
const selectorParser = require('postcss-selector-parser');
const { resolveUiProfileStyleSubset } = require('@spfx-kit/ui-profile/delivery');
const root = path.resolve(__dirname, '..');
const hash = value => crypto.createHash('sha256').update(value).digest('hex');
const files = directory => fs.readdirSync(directory, { withFileTypes: true }).flatMap(item => item.isDirectory() ? files(path.join(directory, item.name)) : [path.join(directory, item.name)]);

function liveKitSource() {
  const imports = new Map(), appSources = [];
  for (const file of files(path.join(root, 'src')).filter(file => /\.tsx?$/.test(file) && !/\.d\.ts$/.test(file))) {
    const source = fs.readFileSync(file, 'utf8'); appSources.push(source);
    const js = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } }).outputText;
    const ast = ts.createSourceFile(file + '.js', js, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
    for (const statement of ast.statements) if (ts.isImportDeclaration(statement) && statement.importClause && statement.moduleSpecifier.text.startsWith('@spfx-kit/')) {
      const clause = statement.importClause.namedBindings;
      if (!clause || !ts.isNamedImports(clause)) throw Error('Review the new non-named Kit import in ' + file);
      const names = clause.elements.map(item => (item.propertyName || item.name).text);
      const module = statement.moduleSpecifier.text;
      if (!imports.has(module)) imports.set(module, new Set());
      for (const name of names) imports.get(module).add(name);
    }
  }
  let index = 0;
  const entry = [...imports].sort(([a], [b]) => a.localeCompare(b)).map(([module, names]) =>
    `export { ${[...names].sort().map(name => `${name} as kit${index++}`).join(', ')} } from ${JSON.stringify(module)};`).join('\n');
  const build = esbuild.buildSync({ stdin: { contents: entry, resolveDir: root }, bundle: true, write: false, treeShaking: true, minifySyntax: true, format: 'esm', platform: 'browser',
    external: ['react', 'react-dom', '@base-ui/*', '@microsoft/*', '@fluentui/*', 'lucide-react', 'motion-dom', 'class-variance-authority', 'clsx', 'tailwind-merge'],
    define: { 'process.env.NODE_ENV': '"production"' } });
  return { entry, sources: [...appSources, build.outputFiles[0].text] };
}

function candidates(sources) {
  // ui-root constructs `style-${name}` at runtime; this app pins the Luma host.
  const values = new Set(['style-luma', 'dark', 'skui:dark']);
  for (const source of sources) {
    const ast = ts.createSourceFile('source.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    const visit = node => {
      if (ts.isStringLiteralLike(node) || ts.isTemplateHead(node) || ts.isTemplateMiddle(node) || ts.isTemplateTail(node)) {
        for (const value of node.text.split(/\s+/)) if (value) values.add(value);
      }
      ts.forEachChild(node, visit);
    };
    visit(ast);
  }
  return values;
}

function possible(selector, values) {
  return selector.nodes.every(node => {
    if (node.type === 'class') return values.has(node.value);
    if (node.type === 'attribute' && node.attribute === 'data-slot' && node.value) {
      return [...values].some(value => node.operator === '^=' ? value.startsWith(node.value) : node.operator === '$=' ? value.endsWith(node.value) : node.operator === '*=' ? value.includes(node.value) : value === node.value);
    }
    // Negative conditions must stay intact; dropping an absent class reverses their meaning.
    if (node.type === 'pseudo' && [':is', ':where', ':has'].includes(node.value)) return node.nodes.some(branch => possible(branch, values));
    return true;
  });
}

function subset(css, values) {
  const tree = postcss.parse(css);
  tree.walkRules(rule => {
    if (rule.parent.type === 'atrule' && /keyframes$/.test(rule.parent.name)) return;
    const selectors = selectorParser().astSync(rule.selector);
    for (const selector of [...selectors.nodes]) if (!possible(selector, values)) selector.remove();
    if (!selectors.nodes.length) rule.remove(); else rule.selector = selectors.toString();
  });
  tree.walkComments(comment => comment.remove());
  // Keep only keyframes referenced by a surviving declaration (including custom properties).
  const declarations = []; tree.walkDecls(declaration => declarations.push(declaration.value));
  tree.walkAtRules(rule => { if (/keyframes$/.test(rule.name) && !declarations.some(value => value.split(/[\s,]+/).includes(rule.params))) rule.remove(); });
  for (let pass = 0; pass < 3; pass++) tree.walkAtRules(rule => { if (rule.nodes && !rule.nodes.length) rule.remove(); });
  return esbuild.transformSync(tree.toString(), { loader: 'css', minifyWhitespace: true }).code.trim();
}

function generate() {
  const live = liveKitSource(), values = candidates(live.sources);
  const sources = {
    profile: resolveUiProfileStyleSubset({ styles: ['luma'] }).bytes.toString('utf8'),
    theme: fs.readFileSync(require.resolve('@spfx-kit/ui-theme/styles.css'), 'utf8'),
    patterns: fs.readFileSync(require.resolve('@spfx-kit/ui-patterns/styles.css'), 'utf8')
  };
  const selected = Object.fromEntries(Object.entries(sources).map(([key, css]) => [key, subset(css, values)]));
  const fonts = postcss.parse(fs.readFileSync(require.resolve('@spfx-kit/ui-theme/fonts/inter.css'), 'utf8'));
  // Preserve international names in normal text and the static English italic setup guidance.
  const fontFiles = ['inter-vietnamese-opsz-normal.woff2', 'inter-latin-ext-opsz-normal.woff2', 'inter-latin-opsz-normal.woff2', 'inter-latin-opsz-italic.woff2'];
  fonts.walkComments(comment => comment.remove());
  fonts.walkAtRules('font-face', face => {
    const props = Object.fromEntries(face.nodes.filter(node => node.type === 'decl').map(node => [node.prop, node.value]));
    if (props.src && props.src.includes('url(') && !fontFiles.some(file => props.src.includes('/' + file))) face.remove();
  });
  const fontCss = fonts.toString().replaceAll('./inter-4.1/', '../../node_modules/@spfx-kit/ui-theme/fonts/inter-4.1/');
  const css = '/* Generated from pinned Kit sources by scripts/subset-kit-styles.cjs. Do not edit. */\n' + Object.values(selected).join('\n') + '\n' + esbuild.transformSync(fontCss, { loader: 'css', minifyWhitespace: true }).code;
  const manifest = { schemaVersion: 1, style: 'luma', cssSha256: hash(css), cssBytes: Buffer.byteLength(css),
    imports: live.entry.split('\n'), sources: Object.fromEntries(Object.entries(sources).map(([key, text]) => [key, { sha256: hash(text), originalBytes: Buffer.byteLength(text), selectedBytes: Buffer.byteLength(selected[key]) }])),
    fonts: fontFiles };
  return { css, manifest, values };
}

if (require.main === module) {
  const { css, manifest } = generate();
  for (const [file, bytes] of Object.entries({ 'src/ui/kit.styles.css': css, 'src/ui/kit.styles.json': JSON.stringify(manifest, null, 2) + '\n' })) {
    if (process.argv.includes('--check')) {
      if (!fs.existsSync(path.join(root,file)) || fs.readFileSync(path.join(root,file), 'utf8').replace(/\r\n/g, '\n') !== bytes) throw Error(file + ' is stale. Run npm run styles:compile.');
    } else fs.writeFileSync(path.join(root,file), bytes);
  }
  console.log('Account Management Kit CSS: ' + manifest.cssBytes + ' bytes; ' + manifest.fonts.length + ' on-demand font subsets; inputs and output verified.');
}
module.exports = { generate, subset, candidates };
