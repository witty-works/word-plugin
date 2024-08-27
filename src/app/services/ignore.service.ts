import { Injectable } from "@angular/core";
import { HttpClient } from "@angular/common/http";
import { environment } from "../../environments/environment";
import { de, en } from "../translations";
import * as Sentry from "@sentry/browser";
import { ErrorUtils } from "../utils/error.utils";

@Injectable({
  providedIn: "root",
})
export class IgnoreService {
  constructor(private http: HttpClient) {}

  async ignoreWordPermanently(
    word: string,
    accessToken: string
  ): Promise<void> {
    const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

    try {
      const requestUrlIgnore = `${environment.dashboard}api/user/language/ignore-words?false_positive=${word}`;

      const httpOptions = {
        headers: {
          Accept: "*/*",
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
        body: {},
      };

      const message = document.getElementById("warn-failed-ignore-error");
      if (message) {
        ErrorUtils.removeErrorMessage(message, "failedRequestText", lang);
      }

      await this.http.put<void>(requestUrlIgnore, {}, httpOptions).toPromise();
    } catch (error: any) {
      Sentry.captureException(error);
       const message = document.getElementById("warn-failed-ignore-error");
      if (message) {
        ErrorUtils.displayErrorMessage(message, "failedRequestText", lang);
      }
      throw error;
    }
  }
}
