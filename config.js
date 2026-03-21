window.__APP_CONFIG__ = window.__APP_CONFIG__ || {};

(function initAppConfig() {
  var host = window.location.hostname || '';
  var isLocal = host === 'localhost' || host === '127.0.0.1' || host === '';
  var isStaging = /staging|preview|dev/i.test(host);
  var env = isLocal ? 'development' : (isStaging ? 'staging' : 'production');

  var defaults = {
    environment: env,
    apiBaseByEnv: {
      development: 'https://mypl-new.mugilvannan.workers.dev',
      staging: '',
      production: 'https://mypl-new.mugilvannan.workers.dev'
    },
    clerkPublishableKeyByEnv: {
      development: 'pk_test_c2V0dGxlZC10b3VjYW4tNTguY2xlcmsuYWNjb3VudHMuZGV2JA',
      staging: '',
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

  merged.apiBase = userConfig.apiBase || merged.apiBaseByEnv[env] || '';
  merged.clerkPublishableKey = userConfig.clerkPublishableKey || merged.clerkPublishableKeyByEnv[env] || '';
  merged.streamEnabled = typeof userConfig.streamEnabled === 'boolean' ? userConfig.streamEnabled : !!merged.streamEnabledByEnv[env];
  merged.streamApiKey = userConfig.streamApiKey || merged.streamApiKeyByEnv[env] || '';
  merged.errorEndpoint = userConfig.errorEndpoint || merged.errorEndpointByEnv[env] || '';

  window.APP_CONFIG = merged;
})();
