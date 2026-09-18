'use strict';
const { spawn } = require('node:child_process');
const path = require('node:path');
const fs = require('node:fs');
const folder = path.resolve(__dirname, '../releases');
fs.mkdirSync(folder, { recursive: true });
if (process.platform === 'win32') spawn('explorer.exe', [folder], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
else console.log(folder);
