import { Injectable } from "@angular/core";
import { HttpClient } from "@angular/common/http";
import { environment } from "../../environments/environment";
import * as Sentry from "@sentry/browser";
import { ErrorUtils } from "../utils/error.utils";
import { getLanguageModule } from "../utils/language.utils";

@Injectable({
  providedIn: "root",
})
export class IgnoreService {
  private lang: any;

  constructor(private http: HttpClient) {
    this.lang = getLanguageModule(); // Initialize the language module once
  }

  async ignoreWordPermanently(
    word: string,
    accessToken: string
  ): Promise<void> {
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
        ErrorUtils.removeErrorMessage(message, "failedRequestText", this.lang);
      }

      await this.http.put<void>(requestUrlIgnore, {}, httpOptions).toPromise();
    } catch (error: any) {
      Sentry.captureException(error);
      const message = document.getElementById("warn-failed-ignore-error");
      if (message) {
        ErrorUtils.displayErrorMessage(message, "failedRequestText", this.lang);
      }
      throw error;
    }
  }
}
