import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { ICheckResponse } from "../data/types";
import { environment } from '../../environments/environment';
import { error } from 'protractor';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  // accessToken = localStorage.getItem('access_token');

  isDevEnv = window.location.hostname === 'localhost'

  constructor(private http: HttpClient) { }

  async makeAuthRequest(): Promise<any> {
    const accessToken = await Office.auth.getAccessToken(); //should always exist

    const url = environment.api + 'v2.0/auth';
    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      }
    };

    console.log('AUTH: url', url, 'httpOptions', httpOptions);
    return await this.http.post<any>(url, {}, httpOptions)
      .toPromise()
      .catch(error => {
        console.log('auth error', error);
        if (error.status === 403) {
          //prompt user to register on dashboard
          const url = 'https://pr-558-6nvmcgq-56xlfiudba6c2.fr-4.platformsh.site' + '/office-register?token=' + accessToken; //TODO
          console.log('url', url)
          Office.context.ui.displayDialogAsync(url, {height: 50, width: 50}, function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log('result.error', result.error);
            }
          });
        }
        throw error;
      }
    ); 
  }

  // makeRefreshTokenRequest(): Promise<any> {
  //   const refreshToken = localStorage.getItem('refresh_token');
  //   const url = environment.dashboard + 'api/refresh-token';

  //   const httpOptions = {
  //     headers: {
  //       Accept: 'application/json',
  //       'Content-Type': 'application/json',
  //     }
  //   };

  //   const body = {
  //     token: refreshToken,
  //   }

  //   return this.http.post<any>(url, body, httpOptions)
  //     .toPromise()
  //     .catch(error => {
  //       // console.error('CORS Error:', error);
  //       // throw error;
  //     });
  // }
}
