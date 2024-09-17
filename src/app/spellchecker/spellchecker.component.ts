import { Component, OnInit, TemplateRef, ViewChild, HostListener } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import { ISpellingError } from "../data/data-structures";
import { IAlternatives, IAlert, IAuthResponse, ICheckResponse, ICheckResponseResult } from "../data/types";
import { AuthService } from '../services/auth.service';
import { environment } from '../../environments/environment';
import DocumentUtils from '../utils/word.utils';
import { useAnalytics } from '../analytics/analytics';
import { DialogRef, DialogService } from "@ngneat/dialog";
import * as Sentry from '@sentry/browser';
import { ErrorUtils } from '../utils/error.utils';
import { KEYBOARD_SHORTCUTS_CONFIG } from '../keyboard-shortcuts.config';
import { getLanguageModule } from '../utils/language.utils';

const analytics = useAnalytics();

/* global Word */
@Component({
  selector: 'app-spellchecker',
  templateUrl: './spellchecker.component.html',
  styleUrls: ['./spellchecker.component.scss']
})
export class SpellcheckerComponent implements OnInit {
  environment = window.location.hostname === 'localhost'
  lang: any;

  isSpellchecking = false;

  hasSpellcheckingRun = false;

  isFirstRun = true;

  isLoggedInWord = true;

  isLoggedin = false;

  isHighlightingCheckedText = false;

  showSpinner = false;

  paragraphsWithIds: { text: string, id: string }[] = [];

  highlights: ISpellingError[] = [];

  lastCorrectedError?: { errorIndex: number, paragraphIndex: number, paragraphText: string, errorText: string };

  hitMaxTextLength = false;

  lastParagraphChecked = 0;

  selectedText = '';

  previouslyCheckedParagraphs: { text: string, id: string, errors: ICheckResponseResult[] }[] = [];

  alerts: IAlert[] = [];
  authResponse: IAuthResponse | null = null;
  checkEndpointResponse: ICheckResponse | null = null;

  @ViewChild('errorDialog') errorDialog?: TemplateRef<any>;


  errorIntro = "";
  errorMessage = "";
  dialogRef?: DialogRef;

  shortcuts = KEYBOARD_SHORTCUTS_CONFIG;

  @HostListener('document:keydown', ['$event'])
  handleKeyboardEvent(event: KeyboardEvent) {
    if (event.ctrlKey && event.shiftKey && event.key === 'C') {
      event.preventDefault();
      this.checkText();
    }
  }

  constructor(
    private spellcheckerService: CheckingService,
    private authService: AuthService,
    private dialogService: DialogService
  ) {
  }

  async ngOnInit() {
    this.lang = getLanguageModule();
    this.register();
    Office.onReady((info) => {
      try {
        Word.run(async (context) => {
          context.document.onParagraphChanged.add(this.paragraphChanged.bind(this));
          await context.sync();
        });
      } catch (error) {
        console.error("Failed to run Word context:", error);
      }
    });
  }

  async paragraphChanged(event: Word.ParagraphChangedEventArgs) {
    return Word.run(async (context) => {
      const body = context.document.body;
      try {
        context.load(body.paragraphs);
        await context.sync();
        const updatedParagraphs = body.paragraphs.load({
          text: true,
        });
        const previousParagraphsWithIds = this.paragraphsWithIds;
        this.paragraphsWithIds = updatedParagraphs.items.map((paragraph) => ({ text: paragraph.text, id: paragraph.uniqueLocalId })).filter(paragraph => paragraph.text !== "");
        const changedParagraph = this.paragraphsWithIds.find(paragraph => paragraph.id === event.uniqueLocalIds[0]);
        if(!changedParagraph) return; 
        await context.sync();

        let offsetChange = 0;
        let indexOfFirstChange = 0
        previousParagraphsWithIds.map((paragraph) => {
          if(paragraph.id === event.uniqueLocalIds[0]) {
            offsetChange = changedParagraph.text.replace(/^\u000b+/, '').length - paragraph.text.replace(/^\u000b+/, '').length;
            //find the first change where previous paragraph and changed paragraph differ
            indexOfFirstChange = paragraph.text.replace(/^\u000b+/, '').split('').findIndex((char, index) => char !== changedParagraph.text.replace(/^\u000b+/, '')[index]);
            return { text: changedParagraph.text, id: paragraph.id };
          }
          return paragraph;
        });

        this.highlights = this.highlights.filter((highlight) => highlight.offset !== indexOfFirstChange);

        //move highlight according to changes
        this.highlights = this.highlights.map((highlight) => {
          if (highlight.paragraphUniqueId === event.uniqueLocalIds[0]) {
            //make sure the highlight is after the change
            if (highlight.offset < indexOfFirstChange) {
              return highlight;
            }
            return {
              ...highlight,
              offset: highlight.offset + offsetChange
            }
          }
          return highlight;
        });
      }
      catch (e) {
        this.handleError(e);
      }
    });
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
      if (response?.code === 13013) { //edge case: throttled
        this.showSpinner = true;
        if (throttleWarning) {
          ErrorUtils.displayErrorMessage(throttleWarning, "throttleWarning", this.lang);
        }
        this.isLoggedin = false;
        localStorage.setItem('is_logged_in', 'false');
        this.hideSpinner();
        return;
      }

      if (throttleWarning) {
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

    Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
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

  async getAccessTokenWithTimestamp(): Promise<{ token: string, timestamp: number }> {
      let accessTokenWithTimestamp = JSON.parse(localStorage.getItem('word_access_token_with_timestamp') ?? '{}');
      const currentTime = new Date().getTime();
    
      // Check if token exists and has more than 20 seconds of lifetime remaining
      if (!accessTokenWithTimestamp?.token || currentTime - accessTokenWithTimestamp.timestamp > environment.tokenLifetime - environment.tokenBuffer) {
        const newAccessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
        });

        // Recalculate the current time to reflect the actual time the new token was fetched
        const newTimestamp = new Date().getTime(); 

        accessTokenWithTimestamp = {
          token: newAccessToken,
          timestamp: newTimestamp // Use new timestamp here
        };
        localStorage.setItem('word_access_token_with_timestamp', JSON.stringify(accessTokenWithTimestamp));
      }
      return accessTokenWithTimestamp;
    }

  async checkText(): Promise<void> {
    this.hitMaxTextLength = false;
    this.isFirstRun = false;
    this.isSpellchecking = true;
    this.selectedText = '';
    this.highlights = [];
    this.previouslyCheckedParagraphs = [];

    try {
      await Word.run(async (context) => {
        // Wrap getSelectedDataAsync in a promise to handle async/await correctly
        const asyncResult = await new Promise<any>((resolve, reject) => {
          Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result);
            } else {
                        reject(new Error('Failed to get selected data: ' + result.error.message));
            }
          });
        });

        this.selectedText = asyncResult.value as string;
        const maxTextLength = environment.maxTextLength;

        if (this.selectedText.length === 0) {
          const body = context.document.body;
                body.load('text');
          await context.sync();
          this.selectedText = body.text.substring(0, maxTextLength);

          if (body.text.length > maxTextLength) {
            this.hitMaxTextLength = true;
          }
        } else if (this.selectedText.length > maxTextLength) {
          this.hitMaxTextLength = true;
          const selectedTextWithinRange = this.selectedText.substring(0, maxTextLength);
                const lastSpace = selectedTextWithinRange.lastIndexOf(' ');
          this.selectedText = this.selectedText.substring(0, lastSpace);
        }

        this.highlights = [];

        // Now process the selected text
        setTimeout(async () => {
          // Continue with further operations inside this callback or call a separate async function
          await this.processSelectedText(context);

          setTimeout(() => {
            this.focusElement("toggle");
          }, 100);
        }, 100);
      });
    } catch (error) {
        console.error('Error in checkText:', error);
    }
  }

  async processSelectedText(context: Word.RequestContext): Promise<void> {
  const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
    const delayDuration = environment.delayDuration;

    try {
      let chunks = this.selectedText.split(/\r/);
    chunks = chunks.filter(paragraph => paragraph.trim() !== ""); // Filter out empty or whitespace-only paragraphs

      const allPageparagraphs = context.document.body.paragraphs;
    allPageparagraphs.load('items');
      await context.sync();

      const paragraphCollection = allPageparagraphs.load({
    select: ['text', 'uniqueLocalId']
      });
      await context.sync();

      this.paragraphsWithIds = paragraphCollection.items
        .filter((paragraph) => paragraph.uniqueLocalId !== null)
        .map((paragraph) => ({
          text: paragraph.text.replace(/\u000b+/g, "").trim(), // Remove all instances of \u000b and trim whitespace
          id: paragraph.uniqueLocalId,
        }))
        .filter((paragraph) => paragraph.text !== "");

      let accessTokenWithTimestamp = await this.getAccessTokenWithTimestamp();

      for (let textChunk of chunks) {
      if (textChunk.trim() === "") continue;  // Skip empty or whitespace-only chunks  
        await delay(delayDuration); // Introduce delay before processing each chunk

        try {
        const newHighlights = await this.spellcheckerService.checkText(textChunk.replace(/^\u000b+/, ''), accessTokenWithTimestamp.token);
          if (!newHighlights) return;

          const newHighlightsExcludingOrthography = {
            ...newHighlights,
            results: newHighlights.results.filter((result: any) => {
            return result.category !== 'orthography' && result.category?.length > 0 && result.subcategory?.length > 0;
          })
          };
          this.checkEndpointResponse = newHighlightsExcludingOrthography;
        if (this.checkEndpointResponse.results.length > 0 && !this.checkEndpointResponse.results[0].alternatives) {  //prompt user to register on dashboard   
          const url = environment.dashboard + 'office-register?token=' + accessTokenWithTimestamp.token;
            Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
              console.log('result.error', result.error);
              }
            });
          }
          const checkLogEventId = Math.random().toString(36).substring(2, 15);
        analytics.checkLog(newHighlightsExcludingOrthography, null, this.selectedText.length, 'check', false, checkLogEventId);

        if (this.authResponse?.plan !== 'witty_free') {
            newHighlightsExcludingOrthography.results.forEach((result: any) => {
            analytics.checkResultLog(result, this.authResponse, this.selectedText.length, 'check_result', false, checkLogEventId);
            });
          }
          //mostly for analytics purposes
          const newAlerts = newHighlightsExcludingOrthography.results.map((result) => ({
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
          }));
          this.alerts = this.alerts.concat(newAlerts);

        newHighlightsExcludingOrthography.results.forEach(highlight => {
            const paragraphIndex = this.paragraphsWithIds.findIndex((paragraph) => {
              //filter out highlights if previous paragraph had identical text -> because we can not differentiate and highlight always the first in this case
            if (this.previouslyCheckedParagraphs.some((checkedParagraph) => checkedParagraph.text === paragraph.text && checkedParagraph.errors.some((error) => error.text === highlight.text))) {
                return false;
              }
              const errorMargin = 0;
            return paragraph.text.substring(highlight.start - errorMargin, highlight.end + errorMargin).includes(highlight.text);
            });
            if (paragraphIndex === -1) return;
            const paragraph = this.paragraphsWithIds[paragraphIndex];
            this.previouslyCheckedParagraphs.push({
              text: paragraph.text,
              id: paragraph.id,
            errors: [highlight]
            });

            this.highlights.push({
              paragraphUniqueId: paragraph.id,
              paragraph: paragraphIndex,
              offset: highlight.start,
              length: highlight.end - highlight.start,
              word: highlight.text,
            details: highlight
            });
          });
          // sort highlights by paragraph -> make sure highlights come in the right order
          this.highlights.sort((a, b) => {
            return a.paragraph - b.paragraph;
          });
          //within the paragraph, order by start offset
          const message = document.getElementById("issue-checking-text");
          if (message) {
            ErrorUtils.removeErrorMessage(message, "issueCheckingText", this.lang);
          }
        } catch (error: any) {
        if (error.name === 'HttpErrorResponse') {
            if (error.status === 422) {
              continue; // Ignore and continue processing the next chunk
            } else if (error.status >= 400 && error.status < 500) {
            Sentry.captureException(new Error(`4xx Error ignored in processSelectedText: ${JSON.stringify(error, null, 2)}`));
              continue;
            }
          } else {
            console.error(error);
          }
        }
      }
    } catch (error: any) {
      Sentry.captureException(new Error(`Error in processSelectedText: ${JSON.stringify(error, null, 2)}`));
      if (error?.code === 13001) {
        this.isLoggedInWord = false;
        this.isLoggedin = false;
      } else {
        const message = document.getElementById("issue-checking-text");
        if (message) {
          ErrorUtils.displayErrorMessage(message, "issueCheckingText", this.lang);
        }
      }
      console.error(error);
    }
    finally {
      this.isSpellchecking = false;
      this.hasSpellcheckingRun = true;
    }
  }

  updatehighlights() {
    this.highlights = this.highlights.map((error, index) => {
      return {
        ...error,
        index: index
      }
    });
  }

  async highlight(obj: { paragraphIndex: number, errorIndex: number }) {
    await Word.run(async (context) => {
      try {
        // Get the paragraph text and the error object using their respective indices.
        const paragraphText = this.paragraphsWithIds[obj.paragraphIndex].text;
        const error = this.highlights[obj.errorIndex];
        // Fetch the paragraph range using the paragraph text.
        const paragraphRange = await DocumentUtils.fetchParagraph(context, paragraphText);

        // Load the paragraph range text to obtain the full content, including any possible changes.
            paragraphRange.load('text');
        await context.sync();

        // Search for all instances of the error word within the paragraph range.
            const searchResults = paragraphRange.search(error.word, { matchCase: true});
            context.load(searchResults, 'text');
        await context.sync();

        // Prepare to find the actual range to highlight by calculating offsets.
        let previousStartOffset = 0;
        let foundMatchingRange = false;
        for (const item of searchResults.items) {
          // Calculate the start offset of this instance of the error word.
          const startOffset = paragraphRange.text.indexOf(item.text, previousStartOffset);
          if (startOffset === error.offset) {
            const errorRange = item;
                    errorRange.select('Select');
            await context.sync();

            foundMatchingRange = true;
            break;
          }
          previousStartOffset = startOffset + 1;
        }

        if (!foundMatchingRange) {
              this.handleError(new Error('The range for the error was not found: ' + error.word));
        }
      } catch (e) {
        this.handleError(e);
      }
    });
  }

  async highlightAndRemoveWordIncludingPreviousSpace(obj: { paragraphIndex: number, errorIndex: number, suggestion: IAlternatives}) {
    await Word.run(async (context) => {
      try {
        const paragraphText = this.paragraphsWithIds[obj.paragraphIndex].text;
        const error = this.highlights[obj.errorIndex];
        const paragraphRange = await DocumentUtils.fetchParagraph(context, paragraphText);

            paragraphRange.load('text');
        await context.sync();

            const searchResults = paragraphRange.search(' ' + error.word, { matchCase: true}); //including previous space
            context.load(searchResults, 'text');
        await context.sync();

        let previousStartOffset = 0;
        let foundMatchingRange = false;
        for (const item of searchResults.items) {
          const startOffset = paragraphRange.text.indexOf(item.text, previousStartOffset) + 1; //+1 accounts for the space
          if (startOffset === error.offset) {
            const errorRange = item;
                    errorRange.select('Select');
            await context.sync();

            foundMatchingRange = true;
            break;
          }
          previousStartOffset = startOffset + 1;
        }

        if (!foundMatchingRange) {
              this.handleError(new Error('The range for the error was not found: ' + error.word));
        }
            const highlightedText = context.document.getSelection()
        highlightedText.load();
        await context.sync();
        highlightedText.insertText(obj.suggestion.text, "Replace");
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
  }

  async acceptSuggestion(obj: { paragraphIndex: number, errorIndex: number, suggestion: IAlternatives }) {
    await Word.run(async (context) => {
      try {
        if(obj.suggestion.remove) {
          this.highlightAndRemoveWordIncludingPreviousSpace(obj);
          return
        }
        const highlightedText = context.document.getSelection()
        highlightedText.load();
        await context.sync();
        highlightedText.insertText(obj.suggestion.text, "Replace");
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
    this.focusElement("toggle");
  }

  insertGrammarError(errorIndex: number, error: ISpellingError) {
    this.highlights.splice(errorIndex, 0, error);
  }

  private handleError(e: any) {

    if (e instanceof Error) {
      if (e.message.startsWith("Could not find range for chunk: ")) {
        this.errorIntro = "Paragraph not found";
        this.errorMessage = e.message.replace("Could not find range for chunk: ", "");
      } else if (e.message.startsWith("The range for the error was not found: ")) {
        this.errorIntro = "Word not found";
        this.errorMessage = e.message.replace("The range for the error was not found: ", "");
      } else {
        this.errorIntro = "Unknown error"
        this.errorMessage = e.message;
      }

      this.dialogRef = this.dialogService.open(this.errorDialog!);
      this.dialogRef.afterClosed$.subscribe((result) => {
        if (!!result) {
          this.checkText();
        }
      });
      console.error(e.message);
    } else {
      console.error(e);
    }
  }

  async highlightCheckedText() {
    this.isHighlightingCheckedText = true;
    await Word.run(async (context) => {
      try {
        if (!this.selectedText || this.selectedText.length <= 255) {
          return;
        }

        let chunks = this.selectedText.split(/\r/);
        chunks = chunks.filter((paragraph) => paragraph !== "" && paragraph !== "\u000b");

        // Prepare all search operations first
        const searchPromises = chunks.map(async (chunk) => {
          if (chunk.length > 200) {
            chunk = chunk.substring(0, 200);
          }
          const searchResults = context.document.body.search(chunk, { matchCase: true, matchWholeWord: false });
                context.load(searchResults, 'items');
          return searchResults; // Return the promise for later resolution
        });

        await context.sync();

        let firstRangeFound = null;
        let lastRangeFound = null;

        // Process search results
        for (const searchPromise of searchPromises) {
          const searchResults = await searchPromise; // Resolve each promise
          const items = searchResults.items;
          if (items.length > 0) {
            if (!firstRangeFound) {
              firstRangeFound = items[0];
            }
            lastRangeFound = items[items.length - 1];
          }
        }

        if (!firstRangeFound || !lastRangeFound) {
          return;
        }

        const completeRange = firstRangeFound.expandTo(lastRangeFound);
            completeRange.select('Select');
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
    this.isHighlightingCheckedText = false;
  }

  focusElement(id: string) {
    setTimeout(() => {
      const elementToFocus = document.getElementById(id);
      if (elementToFocus) {
        elementToFocus.focus();
      }
    }, 100);
  }
}
