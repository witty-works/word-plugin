import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { IAlternatives, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { environment } from '../../environments/environment';
import { de, en } from '../translations';
import * as Sentry from '@sentry/browser';
@Injectable({
  providedIn: 'root'
})
export class CheckingService {
  isDevEnv = window.location.hostname === 'localhost'

  constructor(private settingsService: SettingsService, private http: HttpClient) { }

  async checkText(sentence: string, accessToken: string): Promise<ICheckResponse> {
    try {
      const url = environment.api + 'v2.3/check';

      const body = {
        text: sentence,
        lang: 'auto',
        client: 'word-plugin:' + environment.package_version,
        config: {
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
      Sentry.captureException(new Error(`Error in checkText: ${error}`));
      const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;
      if (error?.code === 13013) { //edge case: throttled
        const throttleWarning = document.getElementById("throttle-warning");
        if (throttleWarning) {
          throttleWarning.style.display = 'block';
          throttleWarning.innerHTML = lang.throttleWarning;
        }
      }

      const errorCodes = [13001, 13002, 13000, 5001];
      if (errorCodes.includes(error?.code)) {
        const message = document.getElementById("warn-not-signed-in-word")

        if (message) {
          message.style.display = 'block';
          message.innerHTML = lang.notSignedInWarning;
        }
      }
      throw error;
    }
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
