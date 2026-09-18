'use strict';

// SPFx's initial bundle already uses classic script loading. Use the same mode
// for lazy JSONP chunks: the managed CDN serves these scripts without CORS
// headers. Setting crossOriginLoading to false omits the crossorigin attribute;
// it does not change the CDN URL, CSP, or SharePoint permissions.
module.exports = config => {
  config.output = { ...config.output, crossOriginLoading: false };
  return config;
};
