import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { environment } from '../../environments/environment';
import * as Sentry from '@sentry/browser';
import { ErrorUtils } from '../utils/error.utils';
import { getLanguageModule } from '../utils/language.utils';
@Injectable({
  providedIn: 'root'
})
export class AuthService {
  isDevEnv = window.location.hostname === 'localhost'
  lang: any;

  constructor(private http: HttpClient) {
    this.lang = getLanguageModule();
  }

  // Check if token is missing, undefined (string), or expired
  private isTokenValid(accessTokenWithTimestamp: any) {
    // Consider checking what a valid token actually looks like (length, can we decode it? can we get a valid email from it?)
    let result = (accessTokenWithTimestamp?.token && accessTokenWithTimestamp.token !== 'undefined');

    if (result) {
      const currentTime = new Date().getTime();
      if ((currentTime - accessTokenWithTimestamp.timestamp) > environment.tokenLifetime - environment.tokenBuffer) {
        result = false;
      } else {
        const remainingTime = environment.tokenBuffer - (currentTime - accessTokenWithTimestamp.timestamp); // Token lifetime in ms

        // Ensure the token has at least 20 seconds of lifetime left before use
        if (remainingTime > environment.tokenBuffer) {
          result = false;
        }
      }
    }

    if (!result) {
      localStorage.removeItem('word_access_token_with_timestamp');
      return false;
    }

    return true;
  }

  async getAccessTokenWithTimestamp(retry = true): Promise<{ token: string, timestamp: number } | null> {
    let accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
    if (!this.isTokenValid(accessTokenWithTimestamp)) {
      if (!retry) {
        return null;
      }

      let result = await this.fetchNewAccessToken();
      if (result) {
        accessTokenWithTimestamp = await this.getAccessTokenWithTimestamp(false)
      }
    }

    ErrorUtils.updateErrorMessages(this.lang, accessTokenWithTimestamp ? null : "warn-not-signed-in-word");

    return accessTokenWithTimestamp;
  }

  async makeAuthRequest(stopTrying: boolean = false): Promise<any> { 
    let accessTokenWithTimestamp = await this.getAccessTokenWithTimestamp()

    try {
      if (!accessTokenWithTimestamp?.token) {
        return;
      }

      const url = environment.api + 'v2.0/auth';
      const httpOptions = {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          Authorization: `Bearer ${accessTokenWithTimestamp.token}`,
        }
      };

      // Making the HTTP POST request
      const response = await this.http.post<any>(url, {}, httpOptions).toPromise();

      return response.plan === null || response.plan === "none" ? "trial-expired" : response;
    } catch (error: any) {
      if (error.status === 403 || error.status === 401) {
        localStorage.removeItem('word_access_token_with_timestamp'); // Clear cached token
        const authFailCounter = localStorage.getItem('authFailCounter') ?? '0';
        if (parseInt(authFailCounter) <= 2) { //has to be 2 to avoid reaching api limit 
          const newCounter = parseInt(authFailCounter) + 1;
          localStorage.setItem('authFailCounter', newCounter.toString());
          setTimeout(() => {
            this.makeAuthRequest(parseInt(authFailCounter) <= 2);
          }, 1000);
        } else if (stopTrying) {
          const url = `${environment.dashboard}office-register?token=${accessTokenWithTimestamp?.token}`;
          if (Office && Office.context && Office.context.ui) {
            Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.log("result.error", result.error);
              }
            });
          }
        }
      }
      return error;
    }
  }

  private async fetchNewAccessToken() {
    if (Office && Office.auth && typeof Office.auth.getAccessToken === 'function') {
      try {
        const newAccessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
        });

        const accessTokenWithTimestamp = {
          token: newAccessToken,
          timestamp: new Date().getTime()
        };

        localStorage.setItem('word_access_token_with_timestamp', JSON.stringify(accessTokenWithTimestamp));
      } catch (error) {
        return false;
      }
    }

    return true;
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
