import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { environment } from '../../environments/environment';
import { de, en } from '../translations';
import * as Sentry from '@sentry/browser';
@Injectable({
  providedIn: 'root'
})
export class AuthService {
  isDevEnv = window.location.hostname === 'localhost'

  constructor(private http: HttpClient) { }

  async makeAuthRequest(): Promise<any> {
    try {
      await Office.onReady();
      let accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');

      if (!accessTokenWithTimestamp?.token || new Date().getTime() - accessTokenWithTimestamp.timestamp > 300000) { //check if token is older than 5 min
           const canGetAccessToken = Office && Office.auth && typeof Office.auth.getAccessToken === 'function';
        if (!canGetAccessToken) {
          throw new Error('Office.auth.getAccessToken is not available');
        }

            const newAccessToken = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
          });

          accessTokenWithTimestamp = {
            token: newAccessToken,
            timestamp: new Date().getTime()
          }
          localStorage.setItem('word_access_token_with_timestamp', JSON.stringify(accessTokenWithTimestamp));
        } 

      const url = environment.api + 'v2.0/auth';
      const httpOptions = {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          Authorization: `Bearer ${accessTokenWithTimestamp.token}`,
        }
      };
      return await this.http.post<any>(url, {}, httpOptions)
        .toPromise()
        .catch(error => {
          if (error.status === 403) {
            localStorage.removeItem('word_access_token_with_timestamp'); //ensures that we get new token
            const authFailCounter = localStorage.getItem('authFailCounter') ?? '0';
            if (parseInt(authFailCounter) <= 2) { //has to be 2 to avoid reaching api limit 
              const newCounter = parseInt(authFailCounter) + 1;
              localStorage.setItem('authFailCounter', newCounter.toString());
              setTimeout(() => {
                this.makeAuthRequest();
              }, 1000);
            } else {
              const url = `${environment.dashboard}office-register?token=${accessTokenWithTimestamp.token}`;
              Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                  console.log('result.error', result.error);
                }
              });
            }
            throw error;
          } else {
            const message = document.getElementById("warn-server-error")
            if (message) {
                const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;
                message.style.display = 'block';
                message.innerHTML = lang.serverError;
            }
          }
        });
    } catch (error: any) {
      const errorMessage = typeof error.message === 'string' ? error.message : JSON.stringify(error, Object.getOwnPropertyNames(error));

      Sentry.captureException(new Error(`Error in makeAuthRequest: ${errorMessage}`));
      const errorCodes = [13001, 13002, 13000, 5001];
      if (errorCodes.includes(error?.code)) {
        const message = document.getElementById("warn-not-signed-in-word")
        if (message) {
            const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;
            message.style.display = 'block';
            message.innerHTML = lang.notSignedInWarning;
        }
      }
      return error;
    }
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
