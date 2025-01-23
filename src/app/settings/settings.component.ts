import { Component, OnDestroy, OnInit } from '@angular/core';
import { SettingsService } from "../services/settings.service";
import { Subscription } from "rxjs";
import { environment } from '../../environments/environment';
import { useAnalytics } from '../analytics/analytics';
import { getLanguageModule } from '../utils/language.utils';

const analytics = useAnalytics();

@Component({
    selector: 'app-settings',
    templateUrl: './settings.component.html',
    styleUrls: ['./settings.component.scss'],
    standalone: false
})
export class SettingsComponent implements OnInit, OnDestroy {

  showContext: boolean = true;
  teamName = '';
  lang: any;
  isLoggedIn = false;
  plan = '';
  public appVersion = '-';

  private showContextSubscription?: Subscription;

  constructor(private settingsService: SettingsService) {
  }

  get isDevEnv(): boolean {
    return window.location.hostname === 'localhost';
  }

  async openDashboard() {
    try {
      analytics.openLinkLog('dashboard_open');
      const accessToken = await Office.auth.getAccessToken(); //can always fetch new here as you will never manage to reach throttle limit
      const url = `${environment.dashboard}office-login?token=${accessToken}`;
      Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log('result.error', result.error);
        }
      });
    } catch (error) {
      throw error;
    }
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
    this.plan = localStorage.getItem('plan') ?? '';
    this.plan = this.plan.replace('_', ' ').replace(/(^\w{1})|(\s+\w{1})/g, letter => letter.toUpperCase());

    this.showContextSubscription = this.settingsService.getShowContextObservable().subscribe(ctx => {
      this.showContext = ctx;
    });
    this.isLoggedIn = localStorage.getItem('is_logged_in') === 'true';
    this.lang = getLanguageModule();
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