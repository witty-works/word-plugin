import { enableProdMode, provideZoneChangeDetection } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';
import './sentry';

if (environment.production) {
  enableProdMode();
}

Office.initialize = () => {
  // Bootstrap the app
  platformBrowserDynamic()
    .bootstrapModule(AppModule, { applicationProviders: [provideZoneChangeDetection()], })
    .catch(error => console.error(error));
};
