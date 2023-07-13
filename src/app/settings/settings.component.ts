import { Component, OnDestroy, OnInit } from '@angular/core';
import { SettingsService } from "../services/settings.service";
import { Subscription } from "rxjs";
// this loads package.json
// then you destructure that object and take out the 'version' property from it
// and finally with ': appVersion' you rename it to const appVersion
const { version: appVersion } = require('../../../package.json');

@Component({
  selector: 'app-settings',
  templateUrl: './settings.component.html',
  styleUrls: ['./settings.component.scss']
})
export class SettingsComponent implements OnInit, OnDestroy {

  showContext: boolean = true;
  checkUpperAndLowerCase: boolean = true;
  checkGrammarAndSpelling: boolean = true;

  public appVersion = '-';

  private showContextSubscription?: Subscription;
  private checkUpperAndLowerCaseSubscription?: Subscription;
  private checkGrammarAndSpellingSubscription?: Subscription;

  constructor(private settingsService: SettingsService) {
  }

  get isDevEnv(): boolean {
    return window.location.hostname === 'localhost'; //TODO: is this the best way to check for dev env?
  }
  
  openDashboard() {
    window.open('https://dashboard.witty.works/en/user/language/customize-witty', '_blank');
  }

  openEditor() {
    window.open('https://dashboard.witty.works/en/editor', '_blank'); 
  }

  openWittyHomePage() {
    window.open('https://witty.works', '_blank');
  }

  ngOnInit() {
    this.appVersion = appVersion;

    this.showContextSubscription = this.settingsService.getShowContextObservable().subscribe(ctx => {
      this.showContext = ctx;
    });

    this.checkUpperAndLowerCaseSubscription = this.settingsService.getCheckUpperAndLowerCaseObservable().subscribe(ctx => {
      this.checkUpperAndLowerCase = ctx;
    });

    this.checkGrammarAndSpellingSubscription = this.settingsService.getCheckGrammarAndSpellingObservable().subscribe(ctx => {
      this.checkGrammarAndSpelling = ctx;
    });
  }

  ngOnDestroy() {
    if (this.showContextSubscription) {
      this.showContextSubscription.unsubscribe();
    }

    if (this.checkUpperAndLowerCaseSubscription) {
      this.checkUpperAndLowerCaseSubscription.unsubscribe();
    }

    if (this.checkGrammarAndSpellingSubscription) {
      this.checkGrammarAndSpellingSubscription.unsubscribe();
    }
  }

  showContextChanged(value: boolean) {
    this.settingsService.setShowContext(value);
  }

  checkUpperAndLowerCaseChanged(value: boolean) {
    this.settingsService.setCheckUpperAndLowerCase(value);
  }

  checkGrammarAndSpellingChanged(value: boolean) {
    this.settingsService.setCheckGrammarAndSpelling(value);
  }
}
