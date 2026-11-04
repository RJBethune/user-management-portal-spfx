'use strict';
const path = require('node:path');
const configureProfile = require('@spfx-kit/ui-profile/spfx-webpack');
const configured = Symbol.for('account-management.ui-kit.subset-configured');
module.exports = function configureKit(config) {
  if (config[configured]) return config;
  for (const rule of config.module?.rules || []) {
    for (const entry of Array.isArray(rule?.use) ? rule.use : [rule?.use]) {
      if (entry && typeof entry === 'object' && typeof entry.loader === 'string' &&
          /(?:^|[\\/])@microsoft[\\/]sp-css-loader(?:[\\/]|$)/.test(entry.loader)) entry.loader = entry.loader.replace(/\\/g, '/');
    }
  }
  // Retain the pinned React 17 compatibility aliases and Luma runtime contract.
  configureProfile(config, { cssDelivery: 'inline', styles: ['luma'] });
  // Our verified component subset replaces the adapter's full Luma delivery asset.
  const marker = Symbol.for('spfx-kit.ui-profile.delivery-assets-plugin');
  if (!config.plugins.some(plugin => plugin?.[marker])) throw Error('Kit delivery adapter changed; review the component subset integration.');
  config.plugins = config.plugins.filter(plugin => !plugin?.[marker]);
  const cssPath = path.normalize(path.resolve(__dirname, '../../lib/ui/kit.styles.css'));
  const matches = resource => path.normalize(String(resource || '').split('?', 1)[0]) === cssPath;
  for (const rule of config.module.rules) rule.exclude = rule.exclude ? [rule.exclude, matches] : matches;
  // Preserve literal React classes and the Kit's own host scope.
  config.module.rules.push({ test: matches, use: [{ loader: path.join(__dirname, 'account-css-loader.cjs') }] });
  Object.defineProperty(config, configured, { value: true });
  return config;
};
