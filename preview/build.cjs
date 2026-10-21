'use strict';
const esbuild = require('esbuild');
const sass = require('sass-embedded');
const fs = require('node:fs');
const path = require('node:path');
import('@spfx-kit/ui-profile/vite').then(({ spfxUiProfileVite }) => esbuild.build({
  alias: spfxUiProfileVite().alias,
  entryPoints: ['preview/entry.tsx'], outdir: 'scratch/preview', bundle: true, platform: 'browser',
  format: 'iife', sourcemap: true, loader: { '.woff2': 'file', '.woff': 'file' },
  define: { 'process.env.NODE_ENV': '"development"' },
  plugins: [{ name: 'fixture-adapters', setup(build) {
    build.onResolve({ filter: /^@microsoft\/sp-(http|core-library|page-context)$/ }, () => ({ path: path.resolve('preview/spfx-stubs.ts') }));
    build.onLoad({ filter: /\.scss$/ }, args => ({ contents: sass.compile(args.path).css, loader: 'local-css', resolveDir: path.dirname(args.path) }));
  } }]
})).then(() => {
  fs.writeFileSync('scratch/preview/index.html', '<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Account Management preview</title><link rel="stylesheet" href="entry.css"><style>body{margin:0;font-family:Segoe UI,sans-serif}#fixture-label{padding:12px;font-size:12px;background:#f0f0f0;color:#333}#app{max-width:1100px;margin:24px auto;padding:0 16px}</style><body><div id="fixture-label">Local UI preview · fictional data · tenant writes disabled</div><div id="app"></div><script src="entry.js"></script></body></html>');
  console.log('Preview built in scratch/preview.');
}).catch(error => { console.error(error); process.exitCode = 1; });
