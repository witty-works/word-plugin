import { Component, ViewChild } from '@angular/core';
import { CheckingService } from "../services/checking.service";
import WordUtils from "../utils/word.utils";
import { ISpellingError } from "../data/data-structures";
import { UserDictionaryService } from "../services/user-dictionary.service";
import { ModalComponent } from "@independer/ng-modal/modal.component";
import { IAlternatives } from "../data/types";

/* global Word */

@Component({
  selector: 'app-spellchecker',
  templateUrl: './spellchecker.component.html',
  styleUrls: ['./spellchecker.component.scss']
})
export class SpellcheckerComponent {

  isSpellchecking = false;

  isFirstRun = true;

  paragraphs: string[] = [];

  spellingErrors: ISpellingError[] = [];

  lastCorrectedError?: { errorIndex: number, paragraphIndex: number, paragraphText: string, errorText: string};

  @ViewChild('errorModal') errorModal?: ModalComponent;

  error = "";

  constructor(
      private spellcheckerService: CheckingService,
      private userDictionaryService: UserDictionaryService,
  ) {
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
          try {
            const errs = await this.spellcheckerService.checkText(paragraph);
            if (!errs) {
              continue;
            }
            console.log(errs);
            errs.results.forEach(e => {
              // TODO: check if in ignore list

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
        const paragraphRange = await WordUtils.getParagraphRange(context, paragraphText);
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
        const paragraphRange = await WordUtils.getParagraphRange(context, paragraphText);
        const errorRange = await WordUtils.getWordRange(context, paragraphRange, errorText);

        errorRange.insertText(obj.suggestion.text, 'Replace');
        errorRange.select('End');

        const newParagraph = paragraphRange.paragraphs.getFirst();
        newParagraph.load('text');
        await context.sync();

        this.updateLineText(obj.paragraphIndex, newParagraph.text);

        // TODO: show toast to revoke change
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
