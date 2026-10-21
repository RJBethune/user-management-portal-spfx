'use strict';
const configureProfile = require('@spfx-kit/ui-profile/spfx-webpack');
const configureTheme = require('@spfx-kit/ui-theme/spfx-webpack.cjs');
module.exports = function configureKit(config) {
  // The released Kit adapter matches the SPFx loader by its POSIX package path.
  for (const rule of config.module?.rules || []) {
    for (const entry of Array.isArray(rule?.use) ? rule.use : [rule?.use]) {
      if (entry && typeof entry === 'object' && typeof entry.loader === 'string' &&
          /(?:^|[\\/])@microsoft[\\/]sp-css-loader(?:[\\/]|$)/.test(entry.loader)) {
        entry.loader = entry.loader.replace(/\\/g, '/');
      }
    }
  }
  return configureTheme(configureProfile(config, { cssDelivery: 'inline', styles: ['luma'] }));
};
