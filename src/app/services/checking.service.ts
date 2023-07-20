import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { BaseUrl, IAlternatives, IBaseUrls, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { AuthService } from './auth.service';

@Injectable({
  providedIn: 'root'
})
export class CheckingService {
  checkUpperAndLowerCase: boolean = true;
  checkGrammarAndSpelling: boolean = true;
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

  isDevEnv = window.location.hostname === 'localhost'

  getBaseUrl(): BaseUrl {
    return this.isDevEnv ? this.BaseUrls.Dev : this.BaseUrls.Prod;
  }

  constructor(private settingsService: SettingsService, private http: HttpClient, private authService: AuthService) {}

  checkText(sentence: string): Promise<ICheckResponse> {
    const accessToken = localStorage.getItem('access_token');
    const refreshToken = localStorage.getItem('refresh_token');
    if ((!accessToken || !sentence) && refreshToken) {
      this.authService.makeRefreshTokenRequest().then((response) => {
        if (response.access_token && response.refresh_token) {
          localStorage.setItem('access_token', response.access_token);
          localStorage.setItem('refresh_token', response.refresh_token);
          this.checkText(sentence);
        } else {
          localStorage.setItem('access_token', '');
          localStorage.setItem('refresh_token', '');
          Promise.resolve({} as ICheckResponse);
        }
      });
    }

    this.settingsService.getCheckUpperAndLowerCaseObservable().subscribe(ctx => {
      this.checkUpperAndLowerCase = ctx;
    });
  
    this.settingsService.getCheckGrammarAndSpellingObservable().subscribe(ctx => {
      this.checkGrammarAndSpelling = ctx;
    });

    const url = this.getBaseUrl().api + 'v2.3/check';

    const body = {
          text: sentence,
          lang: 'auto',
          id: 'appID', //TODO
          client: 'appClient', //TODO
          config: { //TODO: check that this is correct!
            disabled_categories: [
            this.checkGrammarAndSpelling ? '' : 'orthography',
            this.checkUpperAndLowerCase ? '' : 'casing',
          ]},
          config_hash: localStorage.getItem('config_hash'), //dont use config_hash and organization_config_hash for now
          organization_config_hash: localStorage.getItem('organization_config_hash'),
    }

    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      }
    };
  
    return this.http.post<any>(url, body, httpOptions)
      .toPromise()
      .catch(error => {
        // console.error('CORS Error:', error);
        throw error;
      });
  }
  
  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
