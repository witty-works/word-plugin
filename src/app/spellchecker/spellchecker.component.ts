import { Component, OnInit, ViewChild } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import WordUtils from "../utils/word.utils";
import { ISpellingError } from "../data/data-structures";
import { ModalComponent } from "@independer/ng-modal/modal.component";
import { IAlternatives } from "../data/types";
import { AuthService } from '../services/auth.service';
import { en, de } from '../translations';
import { environment } from '../../environments/environment';

/* global Word */

@Component({
  selector: 'app-spellchecker',
  templateUrl: './spellchecker.component.html',
  styleUrls: ['./spellchecker.component.scss']
})
export class SpellcheckerComponent implements OnInit {
  accessToken: string = '';
  refreshToken: string = '';
  environment = window.location.hostname === 'localhost'

  isSpellchecking = false;

  isFirstRun = true;

  isLoggedin = true;

  lang = Office.context.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

  paragraphs: string[] = [];

  spellingErrors: ISpellingError[] = [];

  lastCorrectedError?: { errorIndex: number, paragraphIndex: number, paragraphText: string, errorText: string};

  @ViewChild('errorModal') errorModal?: ModalComponent;

  error = "";

  constructor(private spellcheckerService: CheckingService, private authService: AuthService) {}

  ngOnInit() {
    const accessToken = localStorage.getItem('access_token') || '';
    this.isLoggedin = !!accessToken;

    window.addEventListener('storage', (event) => {
      if (event.key === 'access_token') {
        this.isLoggedin = !!accessToken;
        window.location.reload();
      }
    });    
    
    //probably not needed -> just do auth if check fails, but good for testing
    this.authService.makeAuthRequest().then((response) => {
      if (!response) return;
      localStorage.setItem('organization_name', response.organization_name);
      localStorage.setItem('organization_config_hash', response.organization_config_hash);
      localStorage.setItem('config_hash', response.config_hash);
    });

  }
  
  login() {
    const url = `${environment.dashboard}browser-login?redirect_uri=${environment.plugin + `app/login/login.component.html`}?target=${environment.dashboard}word-addin`;

    Office.context.ui.displayDialogAsync(url, {height: 50, width: 50}, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log('result.error', result.error);
      }
    });
  }

  logout() { 
    localStorage.setItem('access_token', '');
    localStorage.setItem('refresh_token', '');
  }

  openDashboard() {
    Office.context.ui.openBrowserWindow('https://dashboard.witty.works/en/user/language/customize-witty');
  }

  openWittyHomePage() {
    Office.context.ui.openBrowserWindow('https://witty.works');
  }

  async checkGrammar(): Promise<void> {
    this.isFirstRun = false;
    this.isSpellchecking = true;

    return Word.run(async (context) => {
      const body = context.document.body;
      try {
        context.load(body.paragraphs);
        await context.sync();
        const paragraphCollection = body.paragraphs.load({
          text: true,
        });
        this.paragraphs = paragraphCollection.items.map((p) => p.text);

        this.spellingErrors = [];
        for (let paragraphIndex = 0; paragraphIndex < this.paragraphs.length; paragraphIndex++) {
          const paragraph = this.paragraphs[paragraphIndex];
          if (!paragraph) {
            continue;
          }
          try {
            const errs = await this.spellcheckerService.checkText(paragraph.replace(/\u000b/g, '\n'));
            if (!errs) {
              continue;
            }
            errs.results.forEach(e => {
              console.log('event', e);
              if(e.text === ' \v') return; //TODO: handle white space typography error in the future
              this.spellingErrors.push({
                paragraph: paragraphIndex,
                offset: e.start,
                length: e.end - e.start,
                word: e.text,
                details: e
              });
            });
          } catch (e: any) {
            if (e.name === 'HttpErrorResponse' && e.status === 422) {
              continue;
            }
            console.error(e);
          }

        }
      } catch (e) {
        // @ts-ignore
        console.error(e.message, e.debugInfo);
      } finally {
        this.isSpellchecking = false;
      }
    });
  }

  async highlight(obj: {paragraphIndex: number, errorIndex: number }) {
    await Word.run(async (context) => {
      try {
        const paragraphText = this.getLineText(obj.paragraphIndex);
        const errorText = this.getGrammarErrorText(obj.errorIndex);
        const paragraphRange = await WordUtils.getParagraphRange(context, paragraphText, obj.paragraphIndex);
        const errorRange = await WordUtils.getWordRange(context, paragraphRange, errorText);

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
        const paragraphRange = await WordUtils.getParagraphRange(context, paragraphText, obj.paragraphIndex);

        const errorRange = await WordUtils.getWordRange(context, paragraphRange, obj.suggestion.text.length == 0 ? errorText + " " : errorText);

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
    this.errorModal!.closed.subscribe(args => {
      if (!!args.result) {
        this.checkGrammar();
      }
    });
    if (e instanceof Error) {
      if (e.message.startsWith("Could not find range for chunk: ")) {
        // this.error = e.message.replace("Could not find range for chunk: ", "..");
      } else if(e.message.startsWith("The range for the error was not found: ")) {
        // this.error = e.message.replace("The range for the error was not found: ", "..");
      } else {
        this.error = e.message;
      }

      this.errorModal!.open();
      console.error(e.message);
    } else {
      console.error(e);
    }
  }
}
