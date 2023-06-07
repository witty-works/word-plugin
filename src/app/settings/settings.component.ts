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

  public appVersion = '-';

  private showContextSubscription?: Subscription;

  constructor(private settingsService: SettingsService) {
  }

  ngOnInit() {
    this.appVersion = appVersion;

    this.showContextSubscription = this.settingsService.getShowContextObservable().subscribe(ctx => {
      this.showContext = ctx;
    });
  }

  ngOnDestroy() {
    if (this.showContextSubscription) {
      this.showContextSubscription.unsubscribe();
    }
  }

  showContextChanged(value: boolean) {
    this.settingsService.setShowContext(value);
  }
}
