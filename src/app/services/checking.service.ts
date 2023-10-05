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
  checkUpperAndLowerCase: boolean = true;
  checkGrammarAndSpelling: boolean = true;

  isDevEnv = window.location.hostname === 'localhost'

  constructor(private settingsService: SettingsService, private http: HttpClient) {}

  async checkText(sentence: string): Promise<ICheckResponse> {
    const accessToken = await Office.auth.getAccessToken();

    this.settingsService.getCheckUpperAndLowerCaseObservable().subscribe(ctx => {
      this.checkUpperAndLowerCase = ctx;
    });
  
    this.settingsService.getCheckGrammarAndSpellingObservable().subscribe(ctx => {
      this.checkGrammarAndSpelling = ctx;
    });

    const url = environment.api + 'v2.3/check';

    const body = {
          text: sentence,
          lang: 'auto',
          client: 'word-plugin:' + environment.package_version,
          config: { //TODO: check that this is correct!
            disabled_categories: [
            this.checkGrammarAndSpelling ? '' : 'orthography',
            this.checkUpperAndLowerCase ? '' : 'casing',
          ]},
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
  }
  
  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
