import { Injectable } from '@angular/core';
import { SettingsService } from "./settings.service";
import { HttpClient, HttpParams } from "@angular/common/http";
import { IAlternatives, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";

@Injectable({
  providedIn: 'root'
})
export class CheckingService {

  private url = 'https://dev-54ta5gq-him65foajgj5c.fr-4.platformsh.site'; //TODO: implement url logic for dev and prod

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
    return this.http.post<any>(this.url + '/debug/check', request, httpOptions).toPromise();
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
