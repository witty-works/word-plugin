import { Injectable } from "@angular/core";
import { ErrorUtils } from '../utils/error.utils';
import { HttpClient, HttpHeaders } from "@angular/common/http";
import { IAlternative, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { environment } from '../../environments/environment';
import * as Sentry from '@sentry/browser';
import { getLanguageModule } from "../utils/language.utils";
import { TxtSentenceNode } from "sentence-splitter";

@Injectable({
  providedIn: "root",
})
export class CheckingService {
  isDevEnv = window.location.hostname === "localhost";
  lang: any;

  constructor(
    private http: HttpClient
  ) {
    this.lang = getLanguageModule();
  }

  async checkText(
    chunk: string,
    accessToken: string
  ): Promise<ICheckResponse> {

    try {
      const url = environment.api + "v2.4/check";

      const body = {
        text: chunk,
        lang: "auto",
        client: "word-plugin:" + environment.package_version,
        config: {
          disabled_categories: ["orthography"],
        },
        config_hash: localStorage.getItem("config_hash"), //dont use config_hash and organization_config_hash for now
        organization_config_hash: localStorage.getItem(
          "organization_config_hash"
        ),
      };

      const httpOptions = {
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
      };

      ErrorUtils.updateErrorMessages(this.lang);

      return this.http
        .post<any>(url, body, httpOptions)
        .toPromise()
        .catch((error) => {
          throw error;
        });
    } catch (error: any) {
      Sentry.captureException(new Error(`Error in checkText: ${error}`));
      throw error;
    }
  }

  getSuggestions(word: ISpellingError): Promise<IAlternative[]> {
    return Promise.resolve(word.details.alternatives);
  }

  async getLLMSuggestion(error: ISpellingError, sentence: TxtSentenceNode, accessToken: string): Promise<any> {
    const url = environment.api + "v1.0/rephrase";
    const httpOptions = {
      headers: new HttpHeaders({
        "Content-Type": "application/json",
        "Authorization": `Bearer ${accessToken}`,
      }),
    };

    // Send a background API call with the new API input
    const apiInput = {
      sentence: sentence.raw,
      text: error.word,
      start: error.details.start - sentence.range[0],
      alternatives: error.details.alternatives.filter(function (alt) { return !alt.remove; }),
      lang: error.details.language || "en"
    };

    try {
      const response = await this.http
        .post<any>(url, apiInput, httpOptions)
        .toPromise();

      return response;
    } catch (error: any) {
      Sentry.captureException(new Error(`Error in getLLMSuggestion: ${error}`));
      throw error;
    }
  }
}
