import { Injectable } from '@angular/core';
import { HttpClient } from "@angular/common/http";
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  // accessToken = localStorage.getItem('access_token');

  isDevEnv = window.location.hostname === 'localhost'

  constructor(private http: HttpClient) { }

  async makeAuthRequest(): Promise<any> {
    try {
      const accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      });

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
          }
        });
      } catch (error: any) {
        if (error.code === 13001 || error.code === 13002 || error.code === 13000 || error.code === 5001) {
          const message = document.getElementById("warn-not-signed-in-word");
          if (message) {
              message.style.display = 'block';
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
