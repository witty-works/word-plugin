// This file is required by karma.conf.js and loads recursively all the .spec and framework files

import 'zone.js/testing';
import { getTestBed } from '@angular/core/testing';
import {
  BrowserDynamicTestingModule,
  platformBrowserDynamicTesting
} from '@angular/platform-browser-dynamic/testing';

// Office.js is injected by the Office host at runtime and is not available under
// Karma. Stub the surface the app touches so components/services can be constructed.
(globalThis as Record<string, unknown>)['Office'] = {
  initialize: () => undefined,
  onReady: () => Promise.resolve({ host: null, platform: null }),
  AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
  CoercionType: { Text: 'text' },
  PlatformType: { PC: 'PC', Mac: 'Mac', OfficeOnline: 'OfficeOnline' },
  auth: {
    getAccessToken: () => Promise.reject(new Error('not available in tests')),
  },
  context: {
    diagnostics: { platform: 'PC' },
    document: {
      getSelectedDataAsync: (_type: unknown, cb: (r: unknown) => void) =>
        cb({ status: 'failed', value: '' }),
    },
    ui: {
      displayDialogAsync: () => undefined,
      openBrowserWindow: () => undefined,
    },
  },
};

// First, initialize the Angular testing environment.
getTestBed().initTestEnvironment(
  BrowserDynamicTestingModule,
  platformBrowserDynamicTesting(), {
    teardown: { destroyAfterEach: false }
}
);
