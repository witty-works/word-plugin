import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { IAlternatives, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class CheckingService {
  isDevEnv = window.location.hostname === 'localhost'

  constructor(private settingsService: SettingsService, private http: HttpClient) { }

  async checkText(sentence: string): Promise<ICheckResponse> {
    const accessToken = await Office.auth.getAccessToken();

    const url = environment.api + 'v2.3/check';

    const body = {
      text: sentence,
      lang: 'auto',
      client: 'word-plugin:' + environment.package_version,
      config: { //TODO: check that this is correct!
        disabled_categories: [
          'orthography',
        ]
      },
      config_hash: localStorage.getItem('config_hash'), //dont use config_hash and organization_config_hash for now
      organization_config_hash: localStorage.getItem('organization_config_hash'),
    }

    const httpOptions = {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
      }
    };

    return this.http.post<any>(url, body, httpOptions)
      .toPromise()
      .catch(error => {
        if (error.status === 403) {
          //prompt user to register on dashboard
          const url = environment.dashboard + 'office-register?token=' + accessToken; //TODO

          Office.context.ui.displayDialogAsync(url, { height: 50, width: 50 }, function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log('result.error', result.error);
            }
          });
        }
        throw error;
      });
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
