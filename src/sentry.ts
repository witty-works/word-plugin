import * as Sentry from "@sentry/angular";
import { environment } from './environments/environment';

// Imported for side effects from main.ts before the app is bootstrapped, so that
// Sentry is capturing by the time Office.initialize fires. This module must not
// bootstrap the app itself — main.ts owns that, once Office is ready.
if (environment.sentry_dsn) {
  Sentry.init({
    environment: environment.environment,
    release: environment.package_version,
    dsn: environment.sentry_dsn,
    integrations: [
      Sentry.browserTracingIntegration(),
    ],
    sampleRate: environment.sentry_sample_rate,
    tracesSampleRate: environment.sentry_traces_sample_rate,
    replaysSessionSampleRate: 0.0,
  });
}
