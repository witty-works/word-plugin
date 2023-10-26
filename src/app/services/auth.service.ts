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
            const url = environment.dashboard + 'office-register?token=' + accessToken; //TODO

            Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.log('result.error', result.error);
              }
            });
          }
          throw error;
        }

        );
      } catch (error) {
        const ieMessage = document.getElementById("ie-warn");
        if (ieMessage) {
          ieMessage.style.display = 'block';
        }
        throw error;
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
