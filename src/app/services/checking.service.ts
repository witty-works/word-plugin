import { Injectable } from "@angular/core";
import { SettingsService } from "./settings.service";
import { HttpClient } from "@angular/common/http";
import { IAlternatives, ICheckResponse } from "../data/types";
import { ISpellingError } from "../data/data-structures";
import { environment } from '../../environments/environment';
import * as Sentry from '@sentry/browser';
import { ErrorUtils } from "../utils/error.utils";
import { getLanguageModule } from "../utils/language.utils";

@Injectable({
  providedIn: "root",
})
export class CheckingService {
  isDevEnv = window.location.hostname === "localhost";
  lang: any;

  constructor(
    private settingsService: SettingsService,
    private http: HttpClient
  ) {
    this.lang = getLanguageModule(); 
  }

  async checkText(
    sentence: string,
    accessToken: string
  ): Promise<ICheckResponse> {

    try {
      const url = environment.api + "v2.4/check";

      const body = {
        text: sentence,
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

      const throttleWarning = document.getElementById("throttle-warning");
      if (throttleWarning) {
        ErrorUtils.removeErrorMessage(throttleWarning, "throttleWarning", this.lang);
      }

       const authErrorMessage = document.getElementById("warn-not-signed-in-word");
      if (authErrorMessage) {
        ErrorUtils.removeErrorMessage(authErrorMessage, "notSignedInWarning", this.lang);
      }

      const unsupportedAccountErrorMessage = document.getElementById("warn-not-supported-account");
      if (unsupportedAccountErrorMessage) {
        ErrorUtils.removeErrorMessage(unsupportedAccountErrorMessage, "notSupportedAccountWarning", this.lang);
      } 
      
      return this.http
        .post<any>(url, body, httpOptions)
        .toPromise()
        .catch((error) => {
          throw error;
        });
    } catch (error: any) {
       Sentry.captureException(new Error(`Error in checkText: ${error}`));
      if (error?.code === 13013) { //edge case: throttled
        const throttleWarning = document.getElementById("throttle-warning");
        if (throttleWarning) {
          ErrorUtils.displayErrorMessage(throttleWarning, "throttleWarning",this.lang);
        }
      }

      const errorCodes = [13001, 13002, 13000, 5001, 13003];
      if (errorCodes.includes(error?.code)) {
        if (error.code === 13003) {
          const message = document.getElementById("warn-not-supported-account");
          if (message) {
            ErrorUtils.displayErrorMessage(message, "notSupportedAccount", this.lang);
          }
        } else {
          const message = document.getElementById("warn-not-signed-in-word");
          if (message) {
            ErrorUtils.displayErrorMessage(message, "notSignedInWarning", this.lang);
          }
        }
      }
      throw error;
    }
  }

  getSuggestions(word: ISpellingError): Promise<IAlternatives[]> {
    return Promise.resolve(word.details.alternatives);
  }
}
