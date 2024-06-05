import { Injectable } from "@angular/core";
import { HttpClient } from "@angular/common/http";
import { environment } from "../../environments/environment";
import * as Sentry from "@sentry/browser";

@Injectable({
  providedIn: "root",
})
export class IgnoreService {
  constructor(private http: HttpClient) {}

  async ignoreWordPermanently(
    word: string,
    accessToken: string
  ): Promise<void> {
    try {
      console.log(
        `IgnoreService: Initiating API request to ignore word permanently - ${word}`
      );
      const requestUrlIgnore = `${environment.dashboard}api/user/language/ignore-words?false_positive=${word}`;
      console.log(`IgnoreService: API request to URL - ${requestUrlIgnore}`);
      const httpOptions = {
        headers: {
          Accept: "*/*",
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
        body: {},
      };
      console.log(
        `IgnoreService: Retrieved Token successfully: ${accessToken}`
      );

      await this.http.put<void>(requestUrlIgnore, {}, httpOptions).toPromise();
      console.log(`IgnoreService: Successfully ignored word - ${word}`);
    } catch (error: any) {
      console.error(
        `IgnoreService: Error ignoring word '${word}' permanently:`,
        error
      );
      Sentry.captureException(error);
      throw error;
    }
  }
}
