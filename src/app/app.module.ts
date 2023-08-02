import { CUSTOM_ELEMENTS_SCHEMA, NgModule, APP_INITIALIZER, ErrorHandler } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { SettingsComponent } from './settings/settings.component';
import { SpellcheckerComponent } from './spellchecker/spellchecker.component';
import { TabsComponent } from './tabs/tabs.component';
import { ErrorsListComponent } from './spellchecker/errors-list/errors-list.component';
import { ErrorComponent } from './spellchecker/error/error.component';
import { HighlightPipe } from './pipes/highlight.pipe';
import { NgxSpinnerModule } from "ngx-spinner";
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";
import { VirtualScrollerModule } from "@iharbeck/ngx-virtual-scroller";
import { FormsModule } from "@angular/forms";
import { HttpClientModule } from "@angular/common/http";
import { ModalModule } from "@independer/ng-modal";
import { Router } from "@angular/router";
import * as Sentry from "@sentry/angular-ivy";

@NgModule({
  declarations: [
    AppComponent,
    SettingsComponent,
    SpellcheckerComponent,
    TabsComponent,
    ErrorsListComponent,
    ErrorComponent,
    HighlightPipe,
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    HttpClientModule,
    NgxSpinnerModule,
    VirtualScrollerModule,
    AppRoutingModule,
    FormsModule,
    ModalModule,
  ],
  providers: [
    {
     provide: ErrorHandler,
     useValue: Sentry.createErrorHandler({
      showDialog: false,
     }),
    },
    {
     provide: Sentry.TraceService,
     deps: [Router],
    },
    {
     provide: APP_INITIALIZER,
     useFactory: () => () => { },
     deps: [Sentry.TraceService],
     multi: true,
    }
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  bootstrap: [AppComponent]
})
export class AppModule { }
