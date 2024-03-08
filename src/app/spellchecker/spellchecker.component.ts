import { Component, OnInit, TemplateRef, ViewChild } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import { ISpellingError } from "../data/data-structures";
import { IAlternatives, IAlert, IAuthResponse, ICheckResponse } from "../data/types";
import { AuthService } from '../services/auth.service';
import { en, de } from '../translations';
import { environment } from '../../environments/environment';
import DocumentUtils from '../utils/word.utils';
import { useAnalytics } from '../analytics/analytics';
import { DialogRef, DialogService } from "@ngneat/dialog";

const analytics = useAnalytics();

/* global Word */
@Component({
  selector: 'app-spellchecker',
  templateUrl: './spellchecker.component.html',
  styleUrls: ['./spellchecker.component.scss']
})
export class SpellcheckerComponent implements OnInit {
  environment = window.location.hostname === 'localhost'

  isSpellchecking = false;

  isFirstRun = true;

  isLoggedInWord = true;

  isLoggedin = false;

  showSpinner = false;
  
  lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

  paragraphs: string[] = [];

  spellingErrors: ISpellingError[] = [];

  lastCorrectedError?: { errorIndex: number, paragraphIndex: number, paragraphText: string, errorText: string};

  maxTextLength = 1000; //adjust as needed

  hitMaxTextLength = false;

  noParagraphsSelected = false;

  alerts: IAlert[] = [];
  authResponse: IAuthResponse | null = null;
  checkEndpointResponse: ICheckResponse | null = null;

  @ViewChild('errorDialog') errorDialog?: TemplateRef<any>;


  errorIntro = "";
  errorMessage = "";
  dialogRef?: DialogRef;

  constructor(
    private spellcheckerService: CheckingService, 
    private authService: AuthService,
    private dialogService: DialogService
    ) {
    }

  async ngOnInit() { 
    this.register();
  }

  hideSpinner() {
    setTimeout(() => {
      this.showSpinner = false;
    }, 1500);
  }

  register() {
    this.authService.makeAuthRequest().then((response) => {
      const throttleWarning = document.getElementById("throttle-warning");
      const errorCodes = [13001, 13002, 13000, 5001];//not logged in word, did not consent to add-in permissions
      if (errorCodes.includes(response?.code)) {
        this.isLoggedInWord = false;
        this.isLoggedin = false;
        localStorage.setItem('is_logged_in', 'false');
        return;
      } 
      if (response === undefined || response?.status === 403) { //could not authenticate dashboard
        this.isLoggedin = false;
        localStorage.setItem('is_logged_in', 'false');
        return;
      }
      if(response?.code === 13013) { //edge case: throttled
        this.showSpinner = true;
        if(throttleWarning) {
          throttleWarning.style.display = 'block';
          throttleWarning.innerHTML = this.lang.throttleWarning;
        }
        this.isLoggedin = false;
        localStorage.setItem('is_logged_in', 'false');
        this.hideSpinner();
        return;
      }

      if(throttleWarning) {
        throttleWarning.style.display = 'none';
      }

      this.authResponse = response;
      this.isLoggedInWord = true;
      this.isLoggedin = true;
      localStorage.setItem('is_logged_in', 'true');
      localStorage.setItem('organization_name', response.organization_name);
      localStorage.setItem('organization_config_hash', response.organization_config_hash);
      localStorage.setItem('config_hash', response.config_hash);
      localStorage.setItem('user_id', response?.id);
      localStorage.setItem('organization_id', response?.organization_id);
    });
  }
  
  login() {
    const url = `${environment.dashboard}browser-login?redirect_uri=${environment.plugin + `app/login/login.component.html`}?target=${environment.dashboard}word-addin`;

    Office.context.ui.displayDialogAsync(url, {height: 80, width: 80}, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log('result.error', result.error);
      }
    });
  }

  // logout() { 
  //   localStorage.setItem('access_token', '');
  //   localStorage.setItem('refresh_token', '');
  // }

  openWittyHomePage() {
    analytics.openLinkLog('homepage_open');
    Office.context.ui.openBrowserWindow('https://witty.works');
  }

  async checkGrammar(): Promise<void> {
    this.hitMaxTextLength = false;
    this.isFirstRun = false;
    this.isSpellchecking = true;
    let textLengthUsed = 0;
    return Word.run(async (context) => {
      let accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
      if (!accessTokenWithTimestamp?.token || new Date().getTime() - accessTokenWithTimestamp.timestamp > 300000) { //check if token is older than 5 min
        const newAccessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
        });

        accessTokenWithTimestamp = {
          token: newAccessToken,
          timestamp: new Date().getTime()
        }
        localStorage.setItem('word_access_token_with_timestamp', JSON.stringify(accessTokenWithTimestamp));
      }
      const currentlySelectedPageparagraphs = context.document.getSelection().paragraphs
      currentlySelectedPageparagraphs.load();
      await context.sync();
  
      if (currentlySelectedPageparagraphs.items.length === 0 || currentlySelectedPageparagraphs.items[0].text.length === 0) {
        this.noParagraphsSelected = true;
        this.isSpellchecking = false;
        return;
      }
      try {
        context.load(currentlySelectedPageparagraphs);
        await context.sync();
        const paragraphCollection = currentlySelectedPageparagraphs.load({
          text: true,
        });
        this.paragraphs = paragraphCollection.items.map((p) => p.text);

        this.spellingErrors = [];
        for (let paragraphIndex = 0; paragraphIndex < this.paragraphs.length; paragraphIndex++) {
          const paragraph = this.paragraphs[paragraphIndex];
          if (!paragraph) {
            continue;
          }
          if (textLengthUsed + paragraph.length > this.maxTextLength) {
            this.hitMaxTextLength = true;
            break;
          }
          try {
            textLengthUsed += paragraph.length;
            const errs = await this.spellcheckerService.checkText(paragraph.replace(/\u000b/g, '\n'), accessTokenWithTimestamp.token);
            if (!errs) {
              continue;
            }
            this.checkEndpointResponse = errs;
            if (this.checkEndpointResponse.results.length > 0 && !this.checkEndpointResponse.results[0].alternatives) {
              //prompt user to register on dashboard
              const url = environment.dashboard + 'office-register?token=' + accessTokenWithTimestamp.token;
      
              Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                  console.log('result.error', result.error);
                }
              });
            }
            const checkLogEventId = Math.random().toString(36).substring(2, 15);

            analytics.checkLog(errs, null, paragraph.length, 'check', false, checkLogEventId);

            if(this.authResponse?.plan !== 'witty_free') {
              const errsWithoutOrthography = {
                ...errs,
                results: errs.results.filter((result: any) => {
                  return result.category !== 'orthography' && result.category?.length > 0 && result.subcategory?.length > 0;
                }),
              };
              errsWithoutOrthography.results.forEach((result: any) => {
                analytics.checkResultLog(
                  result,
                  this.authResponse,
                  paragraph.length,
                  'check_result',
                  false,
                  checkLogEventId,
                )
              });
            }

            const newAlerts = errs.results.map((result) => ({
              id: `${result.text}-${result.category}-${result.start}${result.end}`,
              startOffset: result.start,
              endOffset: result.end,
              popOverIsOpen: false,
              organizationId: this.authResponse?.organization_id,
              userId: this.authResponse?.id,
              plan: this.authResponse?.plan,
              data: {
                language: this.checkEndpointResponse?.language || 'en',
                category: result.category,
                subcategory: result.subcategory,
                context: result.context,
                text: result.text,
                label: result.label,
                explanation: result.explanation,
                alternatives: result.alternatives,
                gravity: result.gravity,
              },
            }))
            this.alerts = this.alerts.concat(newAlerts);

            errs.results.forEach(e => {
              if(e.text === ' \v') return; //TODO: handle white space typography error in the future
              this.spellingErrors.push({
                paragraph: paragraphIndex,
                offset: e.start,
                length: e.end - e.start,
                word: e.text,
                details: e
              });
            });
          } catch (error: any) {
            if (error.name === 'HttpErrorResponse' && error.status === 422) {
              continue;
            } else if (error?.code === 13001) {
              this.isLoggedInWord = false;
              this.isLoggedin = false;
            }
            console.error(error);
          }

        }
      } catch (error) {
        this.handleError(error);
      } finally {
        this.isSpellchecking = false;
      }
    });
  }

  updateSpellingErrors() {
    this.spellingErrors = this.spellingErrors.map((error, index) => {
      return {
        ...error,
        index: index
      }
    });
  }
  
  async highlight(obj: {paragraphIndex: number, errorIndex: number }) {
    await Word.run(async (context) => {
      try {
        const paragraphText = this.getLineText(obj.paragraphIndex);
        const errorText = this.getGrammarErrorText(obj.errorIndex);
        const paragraphRange = await DocumentUtils.fetchParagraph(context, paragraphText);
        const errorRange = await DocumentUtils.fetchTextBounds(context, paragraphRange, errorText);

        errorRange.select('Select');
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
  }

  acceptSuggestion(obj: {paragraphIndex: number, errorIndex: number, suggestion: IAlternatives }) {
    Word.run(async (context) => {
      try {
        const paragraphText = this.getLineText(obj.paragraphIndex);
        const errorText = this.getGrammarErrorText(obj.errorIndex);
        const paragraphRange = await DocumentUtils.fetchParagraph(context, paragraphText);

        const errorRange = await DocumentUtils.fetchTextBounds(context, paragraphRange, obj.suggestion.text.length == 0 ? errorText + " " : errorText);

        const alertRelevantToSuggestion = this.alerts.find(a => a.data.text === errorText);
        alertRelevantToSuggestion && analytics.alternativeLog(alertRelevantToSuggestion, obj.suggestion.text);

        errorRange.insertText(obj.suggestion.text, 'Replace');        
    
        errorRange.select('End');

        const newParagraph = paragraphRange.paragraphs.getFirst();
        newParagraph.load('text');
        await context.sync();

        this.updateLineText(obj.paragraphIndex, newParagraph.text);

        this.lastCorrectedError = {
          errorIndex: obj.errorIndex,
          paragraphIndex: obj.paragraphIndex,
          paragraphText: paragraphText,
          errorText: errorText
        };

        this.removeGrammarError(obj.errorIndex);

        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
  }

  private getLineText(lineIndex: number): string {
    return this.paragraphs[lineIndex];
  }

  private updateLineText(lineIndex: number, newText: string): void {
    this.paragraphs[lineIndex] = newText;
  }

  private getGrammarErrorText(errorIndex: number): string {
    return this.spellingErrors[errorIndex].word;
  }

  private removeGrammarError(errorIndex: number) {
    this.spellingErrors.splice(errorIndex, 1);
  }

  insertGrammarError(errorIndex: number, error: ISpellingError) {
    this.spellingErrors.splice(errorIndex, 0, error);
  }

  private handleError(e: any) {

    if (e instanceof Error) {
      if (e.message.startsWith("Could not find range for chunk: ")) {
        this.errorIntro = "Betg chattà il paragraf";
        this.errorMessage = e.message.replace("Could not find range for chunk: ", "");
      } else if(e.message.startsWith("The range for the error was not found: ")) {
        this.errorIntro = "Betg chattà il pled";
        this.errorMessage = e.message.replace("The range for the error was not found: ", "");
      } else {
        this.errorIntro = "Errur nunenconuschenta"
        this.errorMessage = e.message;
      }

      this.dialogRef = this.dialogService.open(this.errorDialog!);
      this.dialogRef.afterClosed$.subscribe((result) => {
        if (!!result) {
          this.checkGrammar();
        }
      });
      console.error(e.message);
    } else {
      console.error(e);
    }
  }
  async markLastSpellingError() {
    const lastParagraph = this.spellingErrors[this.spellingErrors.length - 1].paragraph;
    const lastParagraphLastWord = this.paragraphs[lastParagraph].split(' ').pop();
    if(!lastParagraphLastWord) return;
    await Word.run(async (context) => {
      try {
        const paragraphRange = await DocumentUtils.fetchParagraph(context, this.paragraphs[lastParagraph]);
        const errorRange = await DocumentUtils.fetchTextBounds(context, paragraphRange, lastParagraphLastWord);

        errorRange.select('Select');
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
  }
}
