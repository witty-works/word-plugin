import { enableProdMode } from "@angular/core";
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import * as Sentry from "@sentry/angular-ivy";
import { environment } from './environments/environment';
import { AppModule } from "./app/app.module";
  
if (environment.sentry_dsn) {
  Sentry.init({
    environment: environment.environment,
    release: environment.package_version,
    dsn: environment.sentry_dsn,
    integrations: [
    new Sentry.BrowserTracing({
    // Set 'tracePropagationTargets' to control for which URLs distributed tracing should be enabled
    tracingOrigins: [environment.sentry_api_url, /^\//],
    routingInstrumentation: Sentry.routingInstrumentation,
    }),
    new Sentry.Replay(),
    ],
    sampleRate: environment.sentry_sample_rate,
    tracesSampleRate: environment.sentry_traces_sample_rate,
    replaysSessionSampleRate: 0.0,
  });

  enableProdMode();
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .then((success) => console.log('Bootstrap success'))
    .catch((err) => console.error(err));
}