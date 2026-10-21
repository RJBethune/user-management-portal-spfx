'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { createThemeArtifact, compileThemeArtifact } = require('@spfx-kit/ui-theme/compiler');
const pkg = require('../package.json');
for (const name of ['ui-profile', 'ui-theme', 'ui-patterns']) {
  const archive = pkg.dependencies['@spfx-kit/' + name].slice(5);
  const expected = fs.readFileSync(archive + '.sha256', 'utf8').trim().split(/\s/)[0];
  const actual = crypto.createHash('sha256').update(fs.readFileSync(archive)).digest('hex');
  if (actual !== expected) throw Error('Kit archive changed: ' + archive);
}
const profile = JSON.parse(fs.readFileSync(path.join(path.dirname(require.resolve('@spfx-kit/ui-profile/spfx-webpack')), 'profile.json'), 'utf8'));
const document = JSON.parse(fs.readFileSync('vendor/ux-kit-svelte.json', 'utf8'));
const artifact = createThemeArtifact(document, {
  profileId: profile.profileId,
  profileDigest: profile.provenanceSha256,
  revision: crypto.createHash('sha256').update(fs.readFileSync('vendor/ux-kit-svelte.json')).digest('hex')
});
const filename = 'src/ui/kit.compiled.json';
const bytes = JSON.stringify(compileThemeArtifact(artifact), null, 2) + '\n';
if (process.argv.includes('--check')) {
  if (fs.readFileSync(filename, 'utf8').replace(/\r\n/g, '\n') !== bytes) throw Error('Compiled Kit theme is stale');
} else {
  fs.mkdirSync(path.dirname(filename), { recursive: true });
  fs.writeFileSync(filename, bytes);
}
console.log('Kit archive hashes and compiled Svelte theme verified.');
