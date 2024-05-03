import { Component, OnInit, TemplateRef, ViewChild } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import { ISpellingError } from "../data/data-structures";
import { IAlternatives, IAlert, IAuthResponse, ICheckResponse, ICheckResponseResult } from "../data/types";
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

  paragraphsWithIds: { text: string, id: string }[] = [];

  highlights: ISpellingError[] = [];

  lastCorrectedError?: { errorIndex: number, paragraphIndex: number, paragraphText: string, errorText: string };

  hitMaxTextLength = false;

  noParagraphsSelected = false;

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

  constructor(
    private spellcheckerService: CheckingService,
    private authService: AuthService,
    private dialogService: DialogService
  ) {
  }

  async ngOnInit() {
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
          throttleWarning.style.display = 'block';
          throttleWarning.innerHTML = this.lang.throttleWarning;
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
    return accessTokenWithTimestamp;
  }

  async checkText(): Promise<void> {
    this.hitMaxTextLength = false;
    this.noParagraphsSelected = false;
    this.isFirstRun = false;
    this.isSpellchecking = true;
    this.selectedText = ''
    this.highlights = [];
    this.previouslyCheckedParagraphs = [];
    
    return Word.run(async (context) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (asyncResult) => {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          console.log('Action failed. Error: ' + asyncResult.error.message);
        } else {
          this.selectedText = asyncResult.value as string;  
          const maxTextLength = environment.maxTextLength;
          if (this.selectedText.length === 0) {
            this.noParagraphsSelected = true;
            this.highlights = [];
            this.isSpellchecking = false;
            return;
          } else if (this.selectedText.length > maxTextLength) {
            this.hitMaxTextLength = true;
            const selectedTextWithinRange = this.selectedText.substring(0, maxTextLength);
            const lastSpace = selectedTextWithinRange.lastIndexOf(' ');
            this.selectedText = this.selectedText.substring(0, lastSpace);
          }
          
          // Continue with further operations inside this callback or call a separate async function
          await this.processSelectedText(context);
        }
      });
    });
  }
  
  async processSelectedText(context: Word.RequestContext ): Promise<void> {
    try {
      if (this.selectedText.length === 0) return;
      let chunks = this.selectedText.split(/\r/);
      chunks = chunks.filter((paragraph) => paragraph !== "");

      const currentlySelectedPageparagraphs = context.document.getSelection().paragraphs
      currentlySelectedPageparagraphs.load();
      await context.sync();
  
      context.load(currentlySelectedPageparagraphs);
      await context.sync();
      const paragraphCollection = currentlySelectedPageparagraphs.load({
        text: true,
      });

      this.paragraphsWithIds = paragraphCollection.items.map((paragraph) => ({ text: paragraph.text.replace(/^\u000b+/, ''), id: paragraph.uniqueLocalId }));
      this.paragraphsWithIds = this.paragraphsWithIds.filter(paragraph => paragraph.text !== "");
      let accessTokenWithTimestamp = await this.getAccessTokenWithTimestamp();
      for (let textChunk of chunks) {
      if(textChunk === "") continue;
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
          analytics.checkResultLog(result, this.authResponse, this.selectedText.length, 'check_result', false,checkLogEventId)
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
          }))
          this.alerts = this.alerts.concat(newAlerts);


          newHighlightsExcludingOrthography.results.forEach(highlight => {
            const paragraphIndex = this.paragraphsWithIds.findIndex((paragraph) => {
              //filter out highlights if previous paragraph had identical text -> because we can not differentiate and highlight always the first in this case
              if(this.previouslyCheckedParagraphs.some((checkedParagraph) => checkedParagraph.text === paragraph.text && checkedParagraph.errors.some((error) => error.text === highlight.text))) {
                return false;
              } 
              const errorMargin = 0;
              return paragraph.text.substring(highlight.start - errorMargin, highlight.end + errorMargin).includes(highlight.text)
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
        }
    } catch (error: any) {
      if (error.name === 'HttpErrorResponse' && error.status === 422) {
        //TODO: handle this
      } else if (error?.code === 13001) {
        this.isLoggedInWord = false;
        this.isLoggedin = false;
      } else {
        const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;
        const message = document.getElementById("issue-checking-text")
        if (message) {
          message.style.display = 'block';
          message.innerHTML = lang.issueCheckingText;
        }
      }
      console.error(error);
    }
    finally {
      this.isSpellchecking = false;
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

  async acceptSuggestion(obj: { paragraphIndex: number, errorIndex: number, suggestion: IAlternatives }) {
    await Word.run(async (context) => {
      try {
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

  //overly complicated function to highlight the text that was checked (mostly because search is limited to 255 characters), look into simplifying
  async highlightCheckedText() {
    await Word.run(async (context) => {
      try {
        if (!this.selectedText || this.selectedText.length <= 255) {
          // console.log('Text is too short or not specified.');
          return;
        }

        const chunkSize = 255; // Maximum size for each search
        const overlap = 50; // Overlap size to ensure continuity
        let startPosition = 0;
        let endPosition = chunkSize;
        let firstRangeFound = null;
        let lastRangeFound = null;

        while (startPosition < this.selectedText.length && !firstRangeFound) {
          const searchChunk = this.selectedText.substring(startPosition, Math.min(endPosition, this.selectedText.length));
          const searchResults = context.document.body.search(searchChunk, { matchCase: true, matchWholeWord: false });
          context.load(searchResults, 'items');
          await context.sync();

          if (searchResults.items.length > 0) {
            firstRangeFound = searchResults.items[0];
            lastRangeFound = searchResults.items[0];
            break; // Found the first chunk, break the loop to proceed with the next step
          }

          // Adjust positions for the next chunk, considering overlap
          startPosition += (chunkSize - overlap);
          endPosition = startPosition + chunkSize;
        }

        if (!firstRangeFound) {
          // console.log('Text not found in the document.');
          return;
        }

        // Now find the rest of the text, ensuring each chunk is in sequence
        while (endPosition < this.selectedText.length && firstRangeFound && lastRangeFound) {
          const nextChunkStart = endPosition - overlap;
          const nextChunkEnd = nextChunkStart + chunkSize;
          const nextSearchChunk = this.selectedText.substring(nextChunkStart, Math.min(nextChunkEnd, this.selectedText.length));

          const nextSearchResults = context.document.body.search(nextSearchChunk, { matchCase: true, matchWholeWord: false });
          context.load(nextSearchResults, 'items');
          await context.sync();

          if (nextSearchResults.items.length > 0) {
            // Assume the first match is the correct continuation
            lastRangeFound = nextSearchResults.items[0];
            endPosition = nextChunkEnd;
          } else {
            // console.log('Could not find the next part of the text.');
            return;
          }
        }

        if (!lastRangeFound) {
          // console.log('Text not found in the document.');
          return;
        }
        // Highlight the entire text from the first to the last found range
        const completeRange = firstRangeFound.expandTo(lastRangeFound);
        completeRange.select('Select');
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
  }
}
