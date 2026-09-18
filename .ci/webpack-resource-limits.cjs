'use strict';
// Limit SPFx's existing worker pool without changing minification behavior.
module.exports = config => {
  const requested = process.env.EO_SPFX_MINIFIER_THREADS || '1';
  if (!/^[1-8]$/.test(requested)) throw Error('EO_SPFX_MINIFIER_THREADS must be 1 through 8.');
  let matched = 0;
  for (const plugin of config.optimization?.minimizer || []) {
    if (plugin?.minifier && typeof plugin.minifier.maxThreads === 'number') {
      plugin.minifier.maxThreads = Number(requested);
      matched++;
    }
  }
  if (config.mode === 'production' && matched !== 1) throw Error('Expected one SPFx minifier; review this patch after a toolchain upgrade.');
  return config;
};
