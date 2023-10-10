import { Component, OnDestroy, OnInit } from '@angular/core';
import { SettingsService } from "../services/settings.service";
import { Subscription } from "rxjs";
import { en, de } from '../translations';
import { environment } from '../../environments/environment';
import { useAnalytics } from '../analytics/analytics';

const analytics = useAnalytics();

@Component({
  selector: 'app-settings',
  templateUrl: './settings.component.html',
  styleUrls: ['./settings.component.scss']
})
export class SettingsComponent implements OnInit, OnDestroy {

  showContext: boolean = true;
  teamName = '';
  lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

  public appVersion = '-';

  private showContextSubscription?: Subscription;

  constructor(private settingsService: SettingsService) {
  }

  get isDevEnv(): boolean {
    return window.location.hostname === 'localhost';
  }
  
  async openDashboard() {
    analytics.openLinkLog('dashboard_open');
    const accessToken = await Office.auth.getAccessToken();
    console.log('accessToken', accessToken);
    const url = `${environment.dashboard}office-login?token=${accessToken}`;
    Office.context.ui.displayDialogAsync(url, {height: 80, width: 80}, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log('result.error', result.error);
      }
    });
  }

  openWittyHomePage() {
    analytics.openLinkLog('homepage_open');
    Office.context.ui.openBrowserWindow('https://witty.works');
  }

  openHelpCenter() {
    analytics.openLinkLog('helpcenter_open');
    Office.context.ui.openBrowserWindow('https://www.witty.works/en/help/wittys-help-center');
  }

  ngOnInit() {
    this.appVersion = environment.package_version;
    
    this.teamName = localStorage.getItem('organization_name') ?? '';

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

  // logout() {
  //   localStorage.setItem('access_token', '');
  //   localStorage.setItem('refresh_token', '');
  // }
}
