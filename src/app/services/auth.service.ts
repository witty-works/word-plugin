import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { BaseUrl, IBaseUrls, ICheckResponse } from "../data/types";

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  accessToken = localStorage.getItem('access_token');

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

  constructor(private http: HttpClient) {}

  makeAuthRequest(): Promise<any> {
    const accessToken = localStorage.getItem('access_token');
    const refreshToken = localStorage.getItem('refresh_token');

    if (!accessToken && refreshToken) {
      this.makeRefreshTokenRequest().then((response) => {
        if (response.access_token && response.refresh_token) {
          localStorage.setItem('access_token', response.access_token);
          localStorage.setItem('refresh_token', response.refresh_token);
          this.makeAuthRequest();
        } else {
          localStorage.setItem('access_token', '');
          localStorage.setItem('refresh_token', '');
          Promise.resolve({} as ICheckResponse);
        }
      });
    }

    const url = this.getBaseUrl().api + 'v2.0/auth';
    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      }
    };
  
    return this.http.post<any>(url, {}, httpOptions)
      .toPromise()
      .catch(error => {
        // console.error('CORS Error:', error);
        // throw error;
      });
  }

  makeRefreshTokenRequest(): Promise<any> {
    const refreshToken = localStorage.getItem('refresh_token');
    const url = this.getBaseUrl().dashboard + 'api/refresh-token';

    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      }
    };

    const body = {
      token: refreshToken,
    }

    return this.http.post<any>(url, body, httpOptions)
      .toPromise()
      .catch(error => {
        // console.error('CORS Error:', error);
        // throw error;
      });
  }
}
