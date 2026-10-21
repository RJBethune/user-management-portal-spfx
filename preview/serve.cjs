'use strict';
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve('scratch/preview');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' };
http.createServer((request, response) => {
  const name = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
  const target = path.resolve(root, '.' + (name === '/' ? '/index.html' : name));
  if (!target.startsWith(root + path.sep) || !fs.existsSync(target)) { response.writeHead(404).end(); return; }
  response.setHeader('Content-Type', types[path.extname(target)] || 'application/octet-stream');
  fs.createReadStream(target).pipe(response);
}).listen(4318, '127.0.0.1', () => console.log('Fixture preview: http://127.0.0.1:4318'));
