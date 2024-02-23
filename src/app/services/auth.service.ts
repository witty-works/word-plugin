import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { environment } from '../../environments/environment';
import { de, en } from '../translations';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  // accessToken = localStorage.getItem('access_token');

  isDevEnv = window.location.hostname === 'localhost'

  constructor(private http: HttpClient) { }

  async makeAuthRequest(): Promise<any> {
    try {
      let accessToken = localStorage.getItem('word_access_token');

      if (!accessToken) {
        accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
        });
        localStorage.setItem('word_access_token', accessToken);
      }

      const url = environment.api + 'v2.0/auth';
      const httpOptions = {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          Authorization: `Bearer ${accessToken}`,
        }
      };
      return await this.http.post<any>(url, {}, httpOptions)
        .toPromise()
        .catch(error => {
          if (error.status === 403) {
            localStorage.removeItem('word_access_token'); //ensures that we get new token
            const authFailCounter = localStorage.getItem('authFailCounter') ?? '0';
            if (parseInt(authFailCounter) <= 2) { //has to be 2 to avoid reaching api limit 
              const newCounter = parseInt(authFailCounter) + 1;
              localStorage.setItem('authFailCounter', newCounter.toString());
              setTimeout(() => {
                this.makeAuthRequest();
              }, 1000);
            } else {
              const url = `${environment.dashboard}office-register?token=${accessToken}`;
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
