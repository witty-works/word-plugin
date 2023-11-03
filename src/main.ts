import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';
import './sentry';

if (environment.production) {
  enableProdMode();
}

Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);

  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(error => {
      console.error('Error bootstrapping the app:', error);
    });
});
