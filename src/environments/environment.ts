// This file can be replaced during build by using the `fileReplacements` array.
// `ng build --prod` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.

export const environment = {
  package_version: "1.3.18-dev",
  production: false,
  environment: "development",
  sentry_dsn: "",
  sentry_api_url: "https://localhost:4200",
  sentry_sample_rate: 0.0,
  sentry_traces_sample_rate: 0.0,
  sentry_package_name: "word-plugin",
  api: "https://dev-54ta5gq-jyeciedibdzvq.fr-4.platformsh.site/",
  dashboard: "https://dev-54ta5gq-56xlfiudba6c2.fr-4.platformsh.site/",
  plugin: "https://localhost:4200/word-plugin/",
};

/*
 * For easier debugging in development mode, you can import the following file
 * to ignore zone related error stack frames such as `zone.run`, `zoneDelegate.invokeTask`.
 *
 * This import should be commented out in production mode because it will have a negative impact
 * on performance if an error is thrown.
 */
import 'zone.js/plugins/zone-error';  // Included with Angular CLI.
