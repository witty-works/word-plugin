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
    try {
      const accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      });

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
          throw error;
        });
    } catch (error: any) {
      console.log(error);
      if (error.code === 13001) {
        const message = document.getElementById("warn-not-signed-in-word");
        if (message) {
            message.style.display = 'block';
        }
      }
      throw error;
    }
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
