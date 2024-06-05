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
      const requestUrlIgnore = `${environment.dashboard}api/user/language/ignore-words?false_positive=${word}`;

      const httpOptions = {
        headers: {
          Accept: "*/*",
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
        body: {},
      };

      await this.http.put<void>(requestUrlIgnore, {}, httpOptions).toPromise();
    } catch (error: any) {
      Sentry.captureException(error);
      throw error;
    }
  }
}
