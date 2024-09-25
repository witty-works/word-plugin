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

  constructor(private http: HttpClient, private ErrorUtils: ErrorUtils) {
    this.lang = getLanguageModule();
  }

  async makeAuthRequest(): Promise<any> {
    let accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
    try {
      const currentTime = new Date().getTime();

      // Check if token exists and has more than 20 seconds of lifetime remaining
      if (!accessTokenWithTimestamp?.token || (currentTime - accessTokenWithTimestamp.timestamp) > environment.tokenLifetime - environment.tokenBuffer) {
        await this.fetchNewAccessToken();
        accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
      }

      const remainingTime = environment.tokenBuffer - (currentTime - accessTokenWithTimestamp.timestamp); // Token lifetime in ms

      // Ensure the token has at least 20 seconds of lifetime left before use
      if (remainingTime <= environment.tokenBuffer) {
        await this.fetchNewAccessToken();
        accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
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

      if (response.plan === null || response.plan === "none") {
        const trialExpiredMessage = document.getElementById("trial-expired-message");
        if (trialExpiredMessage) {
          ErrorUtils.displayErrorMessage(trialExpiredMessage, "trialExpired", this.lang);
        }
      }

      const authErrorMessage = document.getElementById("warn-not-signed-in-word");
      if (authErrorMessage) {
        ErrorUtils.removeErrorMessage(authErrorMessage, "notSignedInWarning", this.lang);
      }

      const serverErrorMessage = document.getElementById("warn-server-error");
      if (serverErrorMessage) {
        ErrorUtils.removeErrorMessage(serverErrorMessage, "serverError", this.lang);
      }

      const unsupportedAccountErrorMessage = document.getElementById("warn-not-supported-account");
      if (unsupportedAccountErrorMessage) {
        ErrorUtils.removeErrorMessage(unsupportedAccountErrorMessage, "notSupportedAccountWarning", this.lang);
      }

      return response;
    } catch (error: any) {
      if (error.status === 403 || error.status === 401) {
        localStorage.removeItem('word_access_token_with_timestamp'); // Clear cached token
        const authFailCounter = localStorage.getItem('authFailCounter') ?? '0';
        if (parseInt(authFailCounter) <= 2) { //has to be 2 to avoid reaching api limit 
          const newCounter = parseInt(authFailCounter) + 1;
          localStorage.setItem('authFailCounter', newCounter.toString());
          setTimeout(() => {
            this.makeAuthRequest();
          }, 1000);
        } else {
          const url = `${environment.dashboard}office-register?token=${accessTokenWithTimestamp.token}`;
          if (Office && Office.context && Office.context.ui) {
            Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.log("result.error", result.error);
              }
            });
          } else {
          }
        }
      } else {
        const errorMessage = typeof error.message === 'string' ? error.message : JSON.stringify(error, Object.getOwnPropertyNames(error));
        Sentry.captureException(new Error(`Error in makeAuthRequest: ${errorMessage}`));
        const errorCodes = [13001, 13002, 13000, 5001, 13003];
        if (errorCodes.includes(error?.code)) {
          if (error.code === 13003) {
            const message = document.getElementById("warn-not-supported-account");
            if (message) {
              ErrorUtils.displayErrorMessage(message, "notSupportedAccount", this.lang);
            }
          } else {
            const message = document.getElementById("warn-not-signed-in-word");
            if (message) {
              ErrorUtils.displayErrorMessage(message, "notSignedInWarning", this.lang);
            }
          }
        } else {
          const message = document.getElementById("warn-server-error");
          if (message) {
            ErrorUtils.displayErrorMessage(message, "serverError", this.lang);
          }
        }
      }
      return error;
    }
  }

  private async fetchNewAccessToken() {
    if (Office && Office.auth && typeof Office.auth.getAccessToken === 'function') {
      const newAccessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
      });

      const accessTokenWithTimestamp = {
        token: newAccessToken,
        timestamp: new Date().getTime()
      };
      localStorage.setItem('word_access_token_with_timestamp', JSON.stringify(accessTokenWithTimestamp));
    } else {
      const message = document.getElementById("warn-not-signed-in-word");
      if (message) {
        ErrorUtils.displayErrorMessage(message, "notSignedInWarning", this.lang);
      }
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
