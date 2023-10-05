import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { ICheckResponse } from "../data/types";
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  accessToken = localStorage.getItem('access_token');

  isDevEnv = window.location.hostname === 'localhost'

  constructor(private http: HttpClient) { }

  async makeAuthRequest(): Promise<any> {
    // const accessToken = localStorage.getItem('access_token');
    const accessToken = await Office.auth.getAccessToken();
    // const refreshToken = localStorage.getItem('refresh_token');

    // if (!accessToken && refreshToken) {
    //   this.makeRefreshTokenRequest().then((response) => {
    //     if (response.access_token && response.refresh_token) {
    //       localStorage.setItem('access_token', response.access_token);
    //       localStorage.setItem('refresh_token', response.refresh_token);
    //       this.makeAuthRequest();
    //     } else {
    //       localStorage.setItem('access_token', '');
    //       localStorage.setItem('refresh_token', '');
    //       Promise.resolve({} as ICheckResponse);
    //     }
    //   });
    // }

    // const url = environment.api + 'v2.0/auth';
    const url = 'https://pr-879-r66nqla-him65foajgj5c.fr-4.platformsh.site' + 'v2.0/auth';
    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      }
    };

    console.log('AUTH: url', url, 'httpOptions', httpOptions);
    try {
      return await this.http.post<any>(url, {}, httpOptions)
        .toPromise();
    } catch (error) { }
  }

  makeRefreshTokenRequest(): Promise<any> {
    const refreshToken = localStorage.getItem('refresh_token');
    const url = environment.dashboard + 'api/refresh-token';

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
