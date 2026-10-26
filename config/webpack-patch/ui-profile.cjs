'use strict';
const path = require('node:path');
const configureProfile = require('@spfx-kit/ui-profile/spfx-webpack');
const configureTheme = require('@spfx-kit/ui-theme/spfx-webpack.cjs');
const patternsRuleMarker = Symbol.for('account-management.ui-patterns.global-css-rule');
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
  configureTheme(configureProfile(config, { cssDelivery: 'inline', styles: ['luma'] }));

  // ui-patterns 0.5.0 has no SPFx adapter. Its React markup uses literal class
  // names, so its owned stylesheet must bypass SPFx's CSS-module name hashing.
  const rules = config.module.rules;
  const cssPath = path.normalize(require.resolve('@spfx-kit/ui-patterns/styles.css'));
  if (rules.some(rule => rule?.[patternsRuleMarker] === cssPath)) return config;
  const usesSpCssLoader = (rule, modules) => (Array.isArray(rule?.use) ? rule.use : [rule?.use]).some(entry =>
    entry && typeof entry === 'object' && /(?:^|[\\/])@microsoft[\\/]sp-css-loader(?:[\\/]|$)/.test(entry.loader || '') &&
    Boolean(entry.options?.generateCssClassName) === modules);
  const moduleRule = rules.find(rule => usesSpCssLoader(rule, true));
  const globalRule = rules.find(rule => usesSpCssLoader(rule, false));
  if (!moduleRule || !globalRule) throw Error('Could not find the SPFx module and global CSS rules for UI Kit patterns.');
  const matches = resource => path.normalize(String(resource || '').split('?', 1)[0]) === cssPath;
  const exclude = rule => { rule.exclude = rule.exclude ? [rule.exclude, matches] : matches; };
  exclude(moduleRule);
  exclude(globalRule);
  rules.splice(rules.indexOf(globalRule) + 1, 0, {
    ...globalRule, test: matches, include: undefined, exclude: undefined, [patternsRuleMarker]: cssPath
  });
  return config;
};
