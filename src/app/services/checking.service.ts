import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { IAlternatives, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { AuthService } from './auth.service';
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class CheckingService {
  isDevEnv = window.location.hostname === 'localhost'

  constructor(private settingsService: SettingsService, private http: HttpClient, private authService: AuthService) {}

  async checkText(sentence: string): Promise<ICheckResponse> {
    // const accessToken = localStorage.getItem('access_token');
    // const refreshToken = localStorage.getItem('refresh_token');
    // if ((!accessToken || !sentence) && refreshToken) {
    //   this.authService.makeRefreshTokenRequest().then((response) => {
    //     if (response.access_token && response.refresh_token) {
    //       localStorage.setItem('access_token', response.access_token);
    //       localStorage.setItem('refresh_token', response.refresh_token);
    //       this.checkText(sentence);
    //     } else {
    //       localStorage.setItem('access_token', '');
    //       localStorage.setItem('refresh_token', '');
    //       Promise.resolve({} as ICheckResponse);
    //     }
    //   });
    // }
    const accessToken = await Office.auth.getAccessToken();

    const url = environment.api + 'v2.3/check';

    const body = {
          text: sentence,
          lang: 'auto',
          client: 'word-plugin:' + environment.package_version,
          config: { //TODO: check that this is correct!
            disabled_categories: [
            'orthography',
          ]},
          config_hash: localStorage.getItem('config_hash'), //dont use config_hash and organization_config_hash for now
          organization_config_hash: localStorage.getItem('organization_config_hash'),
    }

    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        // Authorization: `Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIzMTE3YmU1YS0zMzIzLTQ1M2YtODEwZC04Nzk1YmFjN2YzMWQiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vN2JkMjJlODQtMzRiMS00YjdiLWI2ZjctNGNkN2RhMGE1ZTRhL3YyLjAiLCJpYXQiOjE2OTQ3ODQ5NjEsIm5i`,
        Authorization: `Bearer ${accessToken}`,
      }
    };
  
    console.log('url', url, 'body', body, 'httpOptions', httpOptions);
    return this.http.post<any>(url, body, httpOptions)
      .toPromise()
      .catch(error => {
        console.log('error', error);
        if (error.status === 403) {
          this.authService.makeAuthRequest().then((response) => {
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
        // console.error('CORS Error:', error);
        throw error;
      });
  }
  
  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
