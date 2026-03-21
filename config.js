window.__APP_CONFIG__ = window.__APP_CONFIG__ || {};

(function initAppConfig() {
  var host = (window.location.hostname || '').toLowerCase();

  var isLocal = host === 'localhost' || host === '127.0.0.1' || host === '';
  var isGitHubPages = host === 'rl-mugil.github.io';
  var isCloudflarePages = host.indexOf('.pages.dev') > -1;
  var isStaging = /staging|preview|dev/i.test(host) && !isCloudflarePages;

  var env = 'production';
  if (isLocal) env = 'development';
  else if (isStaging) env = 'staging';

  var sharedApiBase = 'https://mypl-new.mugilvannan.workers.dev';
  var sharedClerkKey = 'pk_test_c2V0dGxlZC10b3VjYW4tNTguY2xlcmsuYWNjb3VudHMuZGV2JA';

  var defaults = {
    environment: env,
    apiBaseByEnv: {
      development: sharedApiBase,
      staging: sharedApiBase,
      production: sharedApiBase
    },
    clerkPublishableKeyByEnv: {
      development: sharedClerkKey,
      staging: sharedClerkKey,
      production: sharedClerkKey
    },
    errorEndpointByEnv: {
      development: '',
      staging: '',
      production: ''
    },
    appName: 'Metayage Portal'
  };

  var userConfig = window.__APP_CONFIG__;
  var merged = Object.assign({}, defaults, userConfig);

  if (isGitHubPages || isCloudflarePages) {
    merged.environment = 'production';
  }

  merged.apiBase = userConfig.apiBase || merged.apiBaseByEnv[merged.environment] || '';
  merged.clerkPublishableKey = userConfig.clerkPublishableKey || merged.clerkPublishableKeyByEnv[merged.environment] || '';
  merged.errorEndpoint = userConfig.errorEndpoint || merged.errorEndpointByEnv[merged.environment] || '';

  window.APP_CONFIG = merged;
})();
