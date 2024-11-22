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

      ErrorUtils.updateErrorMessages(this.lang);

      await this.http.put<void>(requestUrlIgnore, {}, httpOptions).toPromise();
    } catch (error: any) {
      Sentry.captureException(error);

      ErrorUtils.updateErrorMessages(this.lang, "warn-failed-ignore-error");

      throw error;
    }
  }
}
