import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient, HttpParams } from "@angular/common/http";
import { BaseUrl, IAlternatives, IBaseUrls, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";

@Injectable({
  providedIn: 'root'
})
export class CheckingService {

  BaseUrls: IBaseUrls = {
    Prod: {
      api: 'https://default.api.witty.works/',
      dashboard: 'https://dashboard.witty.works/',
    },
    Dev: {
      api: 'https://dev-54ta5gq-him65foajgj5c.fr-4.platformsh.site/',
      dashboard: 'https://dev-54ta5gq-56xlfiudba6c2.fr-4.platformsh.site/',
    },
  };

  isDevEnv = window.location.hostname === 'localhost'

  getBaseUrl(): BaseUrl {
    return this.isDevEnv ? this.BaseUrls.Dev : this.BaseUrls.Prod;
  }

  constructor(private settingsService: SettingsService, private http: HttpClient) {}

  checkText(sentence: string): Promise<ICheckResponse> {
    const request = {
      type: 'check',
      text: sentence
    };

    let params: HttpParams = new HttpParams();

    const httpOptions = {
      params: params
    };
    return this.http.post<any>(this.getBaseUrl().dashboard + '/debug/check', request, httpOptions).toPromise();
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
