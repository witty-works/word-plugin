import { Component, OnInit, TemplateRef, ViewChild, HostListener } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import { ISpellingError } from "../data/data-structures";
import { IAlternatives, IAlert, IAuthResponse, ICheckResponse } from "../data/types";
import { AuthService } from '../services/auth.service';
import { environment } from '../../environments/environment';
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
  lang: any;

  isSpellchecking = false;

  hasSpellcheckingRun = false;

  isFirstRun = true;

  isLoggedInWord = true;

  isLoggedin = false;

  trialExpired = false;

  isHighlightingCheckedText = false;

  showSpinner = false;

  paragraphsByUniqueId: Map<string, string> = new Map<string, string>();

  highlights: Map<string, ISpellingError[]> = new Map<string, ISpellingError[]>();

  hiddenHighlights: string[] = [];

  ignoredHighlights: string[] = [];

  hitMaxTextLength = false;

  maxTextLength = environment.maxTextLength;

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

  get highlightsList() {
    let highlightList: ISpellingError[] = [];
    for (let key of this.highlights.keys()) {
      let highlights = this.highlights.get(key);
      if (Array.isArray(highlights)) {
        highlights = highlights.filter((highlight) => (
          !this.hiddenHighlights.includes(highlight.errorUniqueId)
          && !this.ignoredHighlights.includes(highlight.errorUniqueId)
        )
        );
        highlightList = highlightList.concat(highlights);
      }
    }

    return highlightList;
  }

  async ngOnInit() {
    this.lang = getLanguageModule();
    this.register();
    Office.onReady((info) => {
      try {
        Word.run(async (context) => {
          context.document.onParagraphChanged.add(this.paragraphChanged.bind(this));
          context.document.onParagraphDeleted.add(this.paragraphDeleted.bind(this));
          await context.sync();
        });
      } catch (error) {
        console.error("Failed to run Word context:", error);
      }
    });
  }

  async paragraphDeleted(event: Word.ParagraphChangedEventArgs) {
    await Word.run(async (context) => {
      let currentParagraphTexts: Map<string, string> = new Map<string, string>();
      let paragraphs = context.document.body.paragraphs;

      paragraphs.load('items');
      await context.sync();

      for (const paragraph of paragraphs.items) {
        paragraph.load("text");
        paragraph.load("uniqueLocalId");
      }
      await context.sync();

      for (const paragraph of paragraphs.items) {
        currentParagraphTexts.set(paragraph.uniqueLocalId, paragraph.text);
      }

      let paragraphArrays = Array.from(this.paragraphsByUniqueId.keys());
      for (const paragraphUniqueId of event.uniqueLocalIds) {
        let paragraphIndex = paragraphArrays.indexOf(paragraphUniqueId);
        if (paragraphIndex < paragraphArrays.length) {
          let paragraphText = this.paragraphsByUniqueId.get(paragraphUniqueId);
          let nextParagraphUniqueId = paragraphArrays[paragraphIndex + 1];
          let nextParagraphText = currentParagraphTexts.get(nextParagraphUniqueId);

          // deleting an empty paragraph can lead to updating the empty paragraph to the content of the previous paragraph (triggers an update before the delete)
          // and deleting the previous paragraph so we need to update the highlights to point to the correct paragraph
          if (nextParagraphText !== undefined && nextParagraphText === paragraphText) {
            let highlights = this.highlights.get(paragraphUniqueId);
            if (highlights !== undefined) {
              for (const highlight of highlights) {
                highlight.paragraphUniqueId = nextParagraphUniqueId;
              }

              let nextHighlights = this.highlights.get(nextParagraphUniqueId);
              if (nextHighlights !== undefined) {
                for (const highlight of nextHighlights) {
                  highlights.push(highlight);
                }
              }
              this.highlights.set(nextParagraphUniqueId, highlights);
            }
          }
        }

        this.highlights.delete(paragraphUniqueId);
        this.paragraphsByUniqueId.delete(paragraphUniqueId);
      }
    });
  }

  async paragraphChanged(event: Word.ParagraphChangedEventArgs) {
    return Word.run(async (context) => {
      try {
        let paragraphs: Map<string, Word.Paragraph> = new Map<string, Word.Paragraph>();

        for (const uniqueLocalIds of event.uniqueLocalIds) {
          let paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(uniqueLocalIds);
          paragraph.load("text")
          paragraphs.set(uniqueLocalIds, paragraph);
        }

        await context.sync();

        Array.from(paragraphs).forEach(([paragraphUniqueId, paragraph]) => {
          this.paragraphsByUniqueId.set(paragraphUniqueId, paragraph.text);

          this.updateParagraphHighlights(paragraphUniqueId);
        });
      } catch (e) {
        this.handleError(e);
      }
    });
  }

  hideSpinner() {
    setTimeout(() => {
      this.showSpinner = false;
    }, 1500);
  }

  isWordEnd(c: string) {
    return !/[-_A-Za-zÀ-ÖØ-öø-ÿ]/.test(c);
  }

  updateParagraphHighlights(paragraphUniqueId: string) {
    let paragraphText = this.paragraphsByUniqueId.get(paragraphUniqueId);
    if (paragraphText === undefined) {
      this.highlights.delete(paragraphUniqueId);
      this.paragraphsByUniqueId.delete(paragraphUniqueId);
      return;
    }

    let highlights = this.highlights.get(paragraphUniqueId);
    if (highlights === undefined || highlights.length == 0) {
      return;
    }

    let highlightsFound: string[] = [];
    let highlightsPositionFound = new Map();

    for (const highlight of highlights) {
      // Start from last position where this specific word was found
      let position = highlightsPositionFound.get(highlight.word);
      let found = false;

      // Try to search for the error across the entire paragraph
      do {
        position = paragraphText.indexOf(highlight.word, position);

        if (position == -1) {
          break;
        }

        let end = position + highlight.word.length

        // check if there is a word end character before
        if (highlight.details.subcategory.indexOf('gendered_denominations_ending') === -1 && position > 0 && !this.isWordEnd(paragraphText[position - 1])) {
          position = end + 1
          continue;
        }

        // check if there is a word end character after
        if (end + 1 < paragraphText.length && !this.isWordEnd(paragraphText[end])) {
          position = end + 1
          continue;
        }

        found = true;
      } while (found === false && position < paragraphText.length)

      // Word not found
      if (!found) {
        continue
      }

      highlightsFound.push(highlight.errorUniqueId)
      highlightsPositionFound.set(highlight.word, position + highlight.word.length + 1);
      highlight.offset = position;

      // Undo hiding if necessary
      // Unfortunately when selecting text and the word appears in the non selected section, it would incorrectly be listed in the sidebar if we do this
      //const errorIndex = this.hiddenHighlights.indexOf(highlight.errorUniqueId);
      //if (errorIndex != -1) {
      //   this.hiddenHighlights.splice(errorIndex, 1);
      //}
    }

    // Hide all errors that have not been found
    for (const highlight of highlights) {
      if (!highlightsFound.includes(highlight.errorUniqueId)
        && !this.hiddenHighlights.includes(highlight.errorUniqueId)
      ) {
        this.hiddenHighlights.push(highlight.errorUniqueId);
      }
    }

    this.highlights.set(paragraphUniqueId, highlights);
  }

  register() {
    this.authService.makeAuthRequest().then((response) => {
      console.log(response);
      if (response === "trial-expired") {
        this.trialExpired = true;
        return;
      }

      if (response?.code) {
        const errorMessage = typeof response.message === 'string' ? response.message : JSON.stringify(response, Object.getOwnPropertyNames(response));
        Sentry.captureException(new Error(`Error in makeAuthRequest: ${errorMessage}`));

        if (response?.code === 13013) { //edge case: unable to get a token
          ErrorUtils.updateErrorMessages(this.lang, "warn-not-supported-account");
          this.isLoggedin = false;
          localStorage.setItem('is_logged_in', 'false');
          return;
        }

        const errorCodes = [13001, 13002, 13000, 5001]; //not logged in word, did not consent to add-in permissions
        if (errorCodes.includes(response?.code)) {
          this.isLoggedInWord = false;
          this.isLoggedin = false;
          localStorage.setItem('is_logged_in', 'false');
          ErrorUtils.updateErrorMessages(this.lang, "warn-not-signed-in-word");
          return;
        } else {
          ErrorUtils.updateErrorMessages(this.lang, "warn-server-error");
          return;
        }
      }

      if (!response || response === undefined || response?.status === 403) { //could not authenticate dashboard
        this.isLoggedin = false;
        localStorage.setItem('is_logged_in', 'false');
        return;
      }

      this.trialExpired = false;

      ErrorUtils.updateErrorMessages(this.lang);

      this.authResponse = response;
      this.isLoggedInWord = true;
      this.isLoggedin = true;
      localStorage.setItem('is_logged_in', 'true');
      localStorage.setItem('organization_name', response.organization_name);
      localStorage.setItem('organization_config_hash', response.organization_config_hash);
      localStorage.setItem('config_hash', response.config_hash);
      localStorage.setItem('user_id', response?.id);
      localStorage.setItem('organization_id', response?.organization_id);
      localStorage.setItem('plan', response.plan);
    });
  }

  login() {
    const url = `${environment.dashboard}browser-login?redirect_uri=${environment.plugin + `app/login/login.component.html`}?target=${environment.dashboard}word-addin`;

    Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('result.error', result.error);
      }
    });
  }

  openWittyHomePage() {
    analytics.openLinkLog('homepage_open');
    Office.context.ui.openBrowserWindow('https://witty.works');
  }

  async openDashboard() {
    try {
      analytics.openLinkLog('dashboard_open');
      const accessToken = await Office.auth.getAccessToken(); //can always fetch new here as you will never manage to reach throttle limit
      const url = `${environment.dashboard}office-login?token=${accessToken}`;
      Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log('result.error', result.error);
        }
      });
    } catch (error) {
      throw error;
    }
  }

  async checkText(): Promise<void> {
    this.hitMaxTextLength = false;
    this.isFirstRun = false;
    this.isSpellchecking = true;
    let selectedText = '';
    this.highlights = new Map();

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

        selectedText = asyncResult.value as string;

        let selectedParagraphs = new Map();
        let paragraphsByUniqueId = new Map();
        let chunks: string[] = [];

        let paragraphs;
        let selection = context.document.getSelection()
        let selectionStart: Word.Range;

        if (selectedText.length === 0) {
          paragraphs = context.document.body.paragraphs;
        } else {
          paragraphs = selection.paragraphs;
          chunks = selectedText.split(/\r/);
        }

        paragraphs.load('items');
        await context.sync();

        for (const paragraph of paragraphs.items) {
          paragraph.load("text");
          paragraph.load("uniqueLocalId");
        }
        await context.sync();

        if (selectedText.length === 0) {
          for (const paragraph of paragraphs.items) {
            chunks.push(paragraph.text);
          }
        }

        let text: string;
        let fullText: string = "";
        for (let i = 0; i < paragraphs.items.length; i++) {
          if (i == 0) {
            text = chunks[0];
          } else if (i == paragraphs.items.length - 1) {
            text = chunks[chunks.length - 1];
          } else {
            text = paragraphs.items[i].text;
          }

          if (fullText.length + text.length >= this.maxTextLength) {
            this.hitMaxTextLength = true;
            const selectedTextWithinRange = text.substring(0, this.maxTextLength - fullText.length);
            const lastSpace = selectedTextWithinRange.lastIndexOf(' ');
            text = text.substring(0, lastSpace);
          }

          selectedParagraphs.set(paragraphs.items[i].uniqueLocalId, text);
          paragraphsByUniqueId.set(paragraphs.items[i].uniqueLocalId, paragraphs.items[i].text);

          if (fullText.length > 0) {
            fullText += "\r"
          }

          fullText += text

          if (this.hitMaxTextLength) {
            break;
          }
        }

        if (fullText.trim() === "") {
          this.isSpellchecking = false;
          ErrorUtils.updateErrorMessages(this.lang, "no-text-selected");
        } else {
          try {
            if (this.hitMaxTextLength) {
              if (selectedText.length === 0) {
                selectionStart = selection.parentBody.getRange("Start");
                selection = selectionStart.expandTo(selection.parentBody.getRange("End"))
              } else {
                selectionStart = selection.getRange('Start');
              }

              const searchResult = await selection.search(fullText.slice(-200), { matchCase: true, matchWholeWord: false });
              context.load(searchResult, 'items');
              await context.sync();

              if (searchResult.items.length) {
                const completeRange = selectionStart.expandTo(searchResult.items[0]);
                completeRange.select('Select');
                await context.sync();
              }
            }
          } catch (error) {
            console.error('Error selecting text:', error);
          }

          this.paragraphsByUniqueId = paragraphsByUniqueId;
          this.highlights = new Map();

          // Now process the selected text
          setTimeout(async () => {
            // Continue with further operations inside this callback or call a separate async function
            await this.processSelectedText(selectedParagraphs);

            setTimeout(() => {
              this.isSpellchecking = false;
              this.focusElement("toggle");
            }, 100);
          }, 100);
        }
      });
    } catch (error) {
      console.error('Error in checkText:', error);
      ErrorUtils.updateErrorMessages(this.lang, "issue-checking-text", `${error}`);
      this.isSpellchecking = false;
    }
  }

  async processSelectedText(selectedParagraphs: Map<string, string>): Promise<void> {
    const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
    const delayDuration = environment.delayDuration;

    this.hasSpellcheckingRun = false;
    let nonFailingCheckResponse = null;
    ErrorUtils.updateErrorMessages(this.lang);

    try {
      let accessTokenWithTimestamp = await this.authService.getAccessTokenWithTimestamp();
      if (!accessTokenWithTimestamp) {
        throw new Error('Valid access token not available');
      }

      let firstParagraph = true;
      for (let paragrapUniqueId of selectedParagraphs.keys()) {
        let selectedParagraph = selectedParagraphs.get(paragrapUniqueId);
        const paragraphText = this.paragraphsByUniqueId.get(paragrapUniqueId);
        // Skip empty or whitespace-only chunks  
        if (selectedParagraph === undefined || paragraphText == undefined || selectedParagraph.trim() === "") {
          continue;
        }

        await delay(delayDuration); // Introduce delay before processing each chunk

        try {
          // The control character could (/^\u000b+/ - vertical tab) exist in the input data (e.g., from a copy-paste operation or a document editor), it standardizes the input text format for further processing.
          const newHighlights = await this.spellcheckerService.checkText(selectedParagraph.replace(/^\u000b+/, ''), accessTokenWithTimestamp.token);
          if (!newHighlights) return;
          this.checkEndpointResponse = newHighlights;
          if (this.checkEndpointResponse.results.length > 0 && !this.checkEndpointResponse.results[0].alternatives) {  //prompt user to register on dashboard   
            const url = environment.dashboard + 'office-register?token=' + accessTokenWithTimestamp.token;
            Office.context.ui.displayDialogAsync(url, { height: 80, width: 80 }, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('result.error', result.error);
              }
            });
          }
          const checkLogEventId = crypto.randomUUID();
          analytics.checkLog(newHighlights, null, selectedParagraph.length, 'check', false, checkLogEventId);

          if (this.authResponse?.plan !== 'witty_free') {
            newHighlights.results.forEach((result: any) => {
              analytics.checkResultLog(result, this.authResponse, selectedParagraph.length, 'check_result', false, checkLogEventId);
            });
          }
          //mostly for analytics purposes
          const newAlerts = newHighlights.results.map((result) => ({
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

          let paragraphOffset = 0;
          if (firstParagraph && selectedParagraph.length < paragraphText.length) {
            paragraphOffset = paragraphText.length - selectedParagraph.length;
            firstParagraph = false;
          }

          newHighlights.results.forEach(highlight => {
            let highlights = this.highlights.get(paragrapUniqueId);
            if (highlights === undefined) {
              highlights = [];
            }

            const error = {
              errorUniqueId: crypto.randomUUID(),
              paragraphUniqueId: paragrapUniqueId,
              offset: highlight.start + paragraphOffset,
              length: highlight.end - highlight.start,
              word: highlight.text,
              details: highlight
            } as ISpellingError;

            highlights.push(error);
            this.highlights.set(paragrapUniqueId, highlights);
          });
          nonFailingCheckResponse = true;
        } catch (error: any) {
          if (error.name === 'HttpErrorResponse') {
            if (error.status === 422) {
              nonFailingCheckResponse = nonFailingCheckResponse === true ? true : '422';
              continue; // Ignore and continue processing the next chunk
            } else if (error.status >= 400 && error.status < 500) {
              Sentry.captureException(new Error(`4xx Error ignored in processSelectedText: ${JSON.stringify(error, null, 2)}`));
              console.error("Status code " + `${error.status}`);
              nonFailingCheckResponse = nonFailingCheckResponse === true ? true : 'xxx';
            } else {
              // in case of f.e. a 500 we hope the next paragraph is ok
              nonFailingCheckResponse = nonFailingCheckResponse === true ? true : 'xxx';
              continue;
            }
          } else {
            nonFailingCheckResponse = nonFailingCheckResponse === true ? true : 'xxx';
            console.error(error);
          }
        }
      }
      this.hasSpellcheckingRun = true;
    } catch (error: any) {
      Sentry.captureException(new Error(`Error in processSelectedText: ${JSON.stringify(error, null, 2)}`));
      ErrorUtils.updateErrorMessages(this.lang, "issue-checking-text", `${error}`);
      console.error(error);
    } finally {
      if (nonFailingCheckResponse === '422') {
        ErrorUtils.updateErrorMessages(this.lang, "cant-identify-language");
      } else if (nonFailingCheckResponse === 'xxx') {
        ErrorUtils.updateErrorMessages(this.lang, "issue-checking-text");
      } else {
        ErrorUtils.updateErrorMessages(this.lang);
      }
    }
  }

  async highlight(obj: { paragraphUniqueId: string, errorUniqueId: string }) {
    await Word.run(async (context) => {
      try {
        // Get the paragraph text and the error object using their respective indices.
        const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(obj.paragraphUniqueId);
        context.load(paragraph, 'text');
        const error = this.getError(obj.paragraphUniqueId, obj.errorUniqueId);

        // Search for all instances of the error word within the paragraph range.
        const searchResults = paragraph.search(error.word, { matchCase: true });
        context.load(searchResults, 'text');
        await context.sync();

        // Prepare to find the actual range to highlight by calculating offsets.
        let previousStartOffset = 0;
        let foundMatchingRange = false;
        for (const item of searchResults.items) {
          // Calculate the start offset of this instance of the error word.
          const startOffset = paragraph.text.indexOf(item.text, previousStartOffset);
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

  getError(paragraphUniqueId: string, errorUniqueId: string): ISpellingError {
    let highlights = this.highlights.get(paragraphUniqueId);
    if (highlights === undefined) {
      throw new Error('Paragraph not found');
    }

    for (const highlight of highlights) {
      if (highlight.errorUniqueId == errorUniqueId) {
        return highlight;
      }
    }

    throw new Error('Error not found in paragraph');
  }

  async highlightAndRemoveWordIncludingPreviousSpace(obj: { paragraphUniqueId: string, errorUniqueId: string, suggestion: IAlternatives }) {
    await Word.run(async (context) => {
      try {
        const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(obj.paragraphUniqueId);
        context.load(paragraph, 'text');
        const error = this.getError(obj.paragraphUniqueId, obj.errorUniqueId);

        const searchResults = paragraph.search(' ' + error.word, { matchCase: true }); //including previous space
        context.load(searchResults, 'text');
        await context.sync();

        let previousStartOffset = 0;
        let foundMatchingRange = false;
        let errorRange;

        for (const item of searchResults.items) {
          const startOffset = paragraph.text.indexOf(item.text, previousStartOffset) + 1; //+1 accounts for the space
          if (startOffset === error.offset) {
            errorRange = item;
            foundMatchingRange = true;
            break;
          }
          previousStartOffset = startOffset + 1;
        }

        if (!foundMatchingRange || !errorRange) {
          this.handleError(new Error('The range for the error was not found: ' + error.word));
          return;
        }

        errorRange.load('text');
        await context.sync();
        errorRange.insertText(obj.suggestion.text, "Replace"); // Directly replace without using document selection
        await context.sync();

      } catch (e) {
        this.handleError(e);
      }
    });
  }

  async acceptSuggestion(obj: { paragraphUniqueId: string, errorUniqueId: string, suggestion: IAlternatives }) {
    await Word.run(async (context) => {
      try {
        if (obj.suggestion.remove) {
          await this.highlightAndRemoveWordIncludingPreviousSpace(obj);
          return;
        }

        // Get the paragraph text and the error object using their respective indices.
        const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(obj.paragraphUniqueId);
        context.load(paragraph, 'text');
        const error = this.getError(obj.paragraphUniqueId, obj.errorUniqueId);

        this.ignoredHighlights.push(error.errorUniqueId);

        const searchResults = paragraph.search(error.word, { matchCase: true });
        context.load(searchResults, 'text');
        await context.sync();

        let previousStartOffset = 0;
        let foundMatchingRange = false;
        let errorRange;

        for (const item of searchResults.items) {
          const startOffset = paragraph.text.indexOf(item.text, previousStartOffset);
          if (startOffset === error.offset) {
            errorRange = item;
            foundMatchingRange = true;
            break;
          }
          previousStartOffset = startOffset + 1;
        }

        if (!foundMatchingRange || !errorRange) {
          this.handleError(new Error('The range for the error was not found: ' + error.word));
          return;
        }

        errorRange.load('text');
        await context.sync();
        errorRange.insertText(obj.suggestion.text, "Replace"); // Directly replace within the found range
        await context.sync();
      } catch (e) {
        this.handleError(e);
      }
    });
    this.focusElement("toggle");
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

      console.error(e.message);

      if (this.errorDialog !== undefined) {
        this.dialogRef = this.dialogService.open(this.errorDialog!);
        this.dialogRef.afterClosed$.subscribe((result) => {
          if (!!result) {
            this.checkText();
          }
        });
      }
    } else {
      console.error(e);
    }
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
