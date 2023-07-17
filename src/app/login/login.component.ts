import { Component, OnDestroy, OnInit } from '@angular/core';
import { BaseUrl, IBaseUrls } from '../data/types';
import { AuthService } from '../auth.service';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})


export class LoginComponent {
  accessToken: string = '';
  refreshToken: string = '';
  isDevEnv = window.location.hostname === 'localhost'

  constructor(private authService: AuthService) {}

  BaseUrls: IBaseUrls = {
    Prod: {
      api: 'https://default.api.witty.works/',
      dashboard: 'https://dashboard.witty.works/',
    },
    Dev: {
      api: 'https://dev-54ta5gq-him65foajgj5c.fr-4.platformsh.site/',
      dashboard: 'https://dev-54ta5gq-56xlfiudba6c2.fr-4.platformsh.site/',
    },
  };

  getBaseUrl(): BaseUrl {
    return this.isDevEnv ? this.BaseUrls.Dev : this.BaseUrls.Prod;
  }
  
  openDashboard() {
    window.open('https://dashboard.witty.works/en/user/language/customize-witty', '_blank');
  }

  openWittyHomePage() {
    window.open('https://witty.works', '_blank');
  }

  login() {
    const url = `${this.getBaseUrl().dashboard}browser-login?redirect_uri=${`https://localhost:4200/word-plugin/app/login/login.component.html`}?target=${this.getBaseUrl().dashboard}editor?onboarding=true`;
    window.open(url, '_blank');
  }

  logout() { 
    this.authService.logout();
  }
}
