window.__APP_CONFIG__ = window.__APP_CONFIG__ || {};

(function initAppConfig() {
  var host = (window.location.hostname || '').toLowerCase();
  var origin = window.location.origin || '';

  var isLocal = host === 'localhost' || host === '127.0.0.1' || host === '';
  var isGitHubPages = host === 'rl-mugil.github.io';
  var isCloudflarePages = host.indexOf('.pages.dev') > -1;
  var isCustomProd = host === 'cms.metayage.com';
  var isWorkersDev = host.indexOf('.workers.dev') > -1;
  var isStaging = /staging|preview/i.test(host);

  var env = 'production';
  if (isLocal) env = 'development';
  else if (isStaging && !isGitHubPages && !isCloudflarePages && !isCustomProd) env = 'staging';

  var defaults = {
    environment: env,
    apiBaseByEnv: {
      development: 'https://mypl-new.mugilvannan.workers.dev',
      staging: 'https://mypl-new.mugilvannan.workers.dev',
      production: 'https://mypl-new.mugilvannan.workers.dev'
    },
    clerkPublishableKeyByEnv: {
      development: 'pk_test_c2V0dGxlZC10b3VjYW4tNTguY2xlcmsuYWNjb3VudHMuZGV2JA',
      staging: 'pk_test_c2V0dGxlZC10b3VjYW4tNTguY2xlcmsuYWNjb3VudHMuZGV2JA',
      production: 'pk_test_c2V0dGxlZC10b3VjYW4tNTguY2xlcmsuYWNjb3VudHMuZGV2JA'
    },
    streamEnabledByEnv: {
      development: true,
      staging: true,
      production: true
    },
    streamApiKeyByEnv: {
      development: '2hhnynjr6ynp',
      staging: '2hhnynjr6ynp',
      production: '2hhnynjr6ynp'
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

  var sameOriginApiBase = '';
  if ((isCustomProd || isWorkersDev) && !isCloudflarePages && !isGitHubPages && origin) {
    sameOriginApiBase = origin;
  }

  merged.apiBase = userConfig.apiBase || sameOriginApiBase || merged.apiBaseByEnv[env] || '';
  merged.clerkPublishableKey = userConfig.clerkPublishableKey || merged.clerkPublishableKeyByEnv[env] || '';
  merged.streamEnabled = typeof userConfig.streamEnabled === 'boolean'
    ? userConfig.streamEnabled
    : !!merged.streamEnabledByEnv[env];
  merged.streamApiKey = userConfig.streamApiKey || merged.streamApiKeyByEnv[env] || '';
  merged.errorEndpoint = userConfig.errorEndpoint || merged.errorEndpointByEnv[env] || '';

  window.APP_CONFIG = merged;
})();
