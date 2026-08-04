import { Component, EventEmitter, Input, Output, HostListener, WritableSignal, signal, ViewEncapsulation, ChangeDetectionStrategy } from "@angular/core";
import { ISpellingError, } from "../../data/data-structures";
import TextUtils from "../../utils/text.utils";
import { CheckingService } from "../../services/checking.service";
import { IgnoreService } from "../../services/ignore.service";
import { AuthService } from '../../services/auth.service';
import { SpellcheckerComponent } from "../spellchecker.component";
import { IAlert, IAlternative, IRephrasingResult, DiffChange } from "../../data/types";
import { useAnalytics } from "../../analytics/analytics";
import * as Sentry from "@sentry/browser";
import { KEYBOARD_SHORTCUTS_CONFIG } from "../../keyboard-shortcuts.config";
import { getLanguageModule } from '../../utils/language.utils';
import { TxtSentenceNode } from 'sentence-splitter';
import { diffWords } from 'diff';

const analytics = useAnalytics();

@Component({
    selector: "app-error",
    templateUrl: "./error.component.html",
    styleUrls: ["./error.component.scss"],
    encapsulation: ViewEncapsulation.None,
    changeDetection: ChangeDetectionStrategy.Eager,
    standalone: false
})
export class ErrorComponent {
  @Input()
  error?: ISpellingError;

  @Input()
  showContext: boolean = true;

  @Output()
  highlightEvent = new EventEmitter();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ suggestion: IAlternative }>();

  shortcuts = KEYBOARD_SHORTCUTS_CONFIG;

  @HostListener('keydown', ['$event'])
  handleKeyboardEvent(event: KeyboardEvent) {
    if (event.ctrlKey && event.shiftKey) {
      switch (event.key) {
        case 'I':
          event.preventDefault();
          this.ignoreOnce();
          break;
        case 'P':
          event.preventDefault();
          if (this.error?.word) {
            this.ignorePermanently(this.error.word);
          }
          break;
      }
    }
  }

  isOpen = false;
  suggestions: IAlternative[] = [];
  showLearningBite: boolean = false;

  alerts: IAlert[] = this.spellcheckerComponent.alerts;

  paragraphsByUniqueId: Map<string, string> = this.spellcheckerComponent.paragraphsByUniqueId;

  highlights: ISpellingError[] = this.spellcheckerComponent.highlightsList;

  lang: any;

  rephrasing: string | null = null;

  suggestion: IAlternative | null = null;

  rephrasings: WritableSignal<IRephrasingResult | null> = signal(null);

  constructor(
    private spellcheckerService: CheckingService,
    private spellcheckerComponent: SpellcheckerComponent,
    private authService: AuthService,
    private ignoreService: IgnoreService
  ) {
    this.lang = getLanguageModule();
  }

  private removeHTMLTags(htmlString: string) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlString, 'text/html');
    const textContent = doc.body.textContent ?? '';
    return textContent.trim();
  }

  private computeDiff(language: string, originalSentence: string, newSentence: string) {
    const options = {
      intlSegmenter: new (Intl as any).Segmenter(language, { granularity: 'word' })
    }

    const diffElements: DiffChange[] = diffWords(
      originalSentence,
      newSentence,
      options
    );

    const maxDiffLength = 120;
    const minDiffLength = 50;

    let diffElement: DiffChange;

    let diff: string = '';
    let tag: string = '';

    let start: number = 0;
    let end: number = diffElements.length - 1;

    for (let i = 0; i < diffElements.length; i++) {
      diffElement = diffElements[i]
      if (diffElement.added || diffElement.removed) {
        if (diff === '') {
          start = i;
        } else if (diff.endsWith('</ins>') || diff.endsWith('</del>')) {
          diff += ' '
        }
        end = i;

        tag = diffElement.added ? 'ins' : 'del';
        diff += `<${tag}>${diffElement.value}</${tag}>`;
      } else if (diff !== '' && i < diffElements.length - 1) {
        diff += diffElement.value;
      }
    }

    const strippedString = this.removeHTMLTags(diff);
    let stringLengthDiff = maxDiffLength - strippedString.length
    if (stringLengthDiff > 0) {
      let value = '';
      if (start > 0) {
        value = diffElements[0].value.substring(
          Math.max(0, diffElements[0].value.length - Math.min(minDiffLength, stringLengthDiff)),
          diffElements[0].value.length
        );

        if (value != diffElements[0].value) {
          value = value.slice(value.indexOf(' '));
          value = '...' + value
        }

        diff = value + diff;
      }

      stringLengthDiff -= value.length
      if (end < diffElements.length - 1 && stringLengthDiff) {
        let value = diffElements[diffElements.length - 1].value.substring(0, Math.min(minDiffLength, stringLengthDiff));
        if (value != diffElements[diffElements.length - 1].value) {
          value = value.substring(0, value.lastIndexOf(' '));
          value = value + '...'
        }

        diff = diff + value;
      }
    }

    return diff
  }

  suggestionedHovered(suggestion: IAlternative, hovered: boolean | null = null) {
    if (hovered !== null) {
      suggestion.hovered = hovered;
    }

    if (!suggestion.hovered || !this.error) {
      this.rephrasing = null;
      this.suggestion = null;
      return;
    }

    this.suggestion = suggestion;

    let rephrasing: string | undefined;
    let sentence: TxtSentenceNode | null | string = null;
    if (suggestion.remove) {
      sentence = this.spellcheckerComponent.getSentence(this.error);
      if (sentence) {
        rephrasing = this.spellcheckerComponent.simpleReplacement(sentence, this.error, '');
        sentence = sentence.raw
      }
    } else {
      const rephrasings: IRephrasingResult | null = this.rephrasings()
      if (rephrasings === null) {
        this.rephrasing = `<img aria-live="polite" role="alert" src="assets/icons/spinner.svg" alt="${this.lang.loading}" />`;
        return;
      }
  
      rephrasing = rephrasings.results.get(suggestion.text);
      if (rephrasing === undefined) {
        sentence = this.spellcheckerComponent.getSentence(this.error);
        if (sentence) {
          rephrasing = this.spellcheckerComponent.simpleReplacement(sentence, this.error, suggestion.text);
          sentence = sentence.raw;
        }
      } else {
        sentence = rephrasings.sentence;
      }
    }

    if (sentence === null || rephrasing === undefined) {
      return;
    }

    this.rephrasing = this.computeDiff(this.error.details.language, sentence, rephrasing);
  }

  getContextErrorComponent(error: ISpellingError) {
    let ctxt = TextUtils.getContext(error, this.paragraphsByUniqueId);
    if (ctxt) {
      ctxt = ctxt.replace(
        /()/g,
        '<img src="assets/icons/soft-return.svg" class="soft-return-icon" alt="Soft return icon"><br>'
      );
    } else {
      this.sendErrorToSentry();
    }
    return ctxt;
  }

  async toggle(): Promise<void> {
    this.suggestions = await this.spellcheckerService.getSuggestions(
      this.error!
    );
    const alertRelevantToSuggestion = this.alerts.find((alert) => {
      return alert.data.text === this.error?.word;
    });

    this.suggestions = this.suggestions.map((suggestion) => {
      if (suggestion && suggestion.text) {
        suggestion.text = suggestion.text
          .replace(/\(\(/g, "[")
          .replace(/\)\)/g, "]");
      }
      return suggestion;
    });

    if (!this.isOpen) {
      this.isOpen = true;
      alertRelevantToSuggestion &&
        analytics.popoverLogs(alertRelevantToSuggestion, "popover_open");
      this.sendHighlight();
    } else {
      this.isOpen = false;
      alertRelevantToSuggestion &&
        analytics.popoverLogs(alertRelevantToSuggestion, "popover_close");
    }

    // TODO handle error undefined
    if (!this.isOpen || this.error === undefined) {
      return;
    }

    const sentence = this.spellcheckerComponent.getSentence(this.error)
    if (sentence === null) {
      // TODO handle error
      return;
    }

    const rephrasings = await this.spellcheckerComponent.fetchRephrasings(this.error, sentence);
    this.error.rephrasings = rephrasings
    this.rephrasings.set(rephrasings)
    if (this.suggestion) {
      this.suggestionedHovered(this.suggestion);
    }
  }

  sendHighlight() {
    this.highlightEvent.emit();
  }

  acceptSuggestion(suggestion: IAlternative) {
    if (suggestion.remove) {
      suggestion.text = "";
    }
    this.acceptSuggestionEvent.emit({ suggestion });
    this.focusElement("toggle");
  }

  onClick(url: string | undefined) {
    this.showLearningBite = !this.showLearningBite;
    const alertRelevantToSuggestion = this.alerts.find((alert) => {
      return alert.data.text === this.error?.word;
    });
    alertRelevantToSuggestion &&
      analytics.popoverLogs(alertRelevantToSuggestion, "learning_bites");
    url && Office.context.ui?.openBrowserWindow(url);
  }

  ignoreOnce() {
    if (this.error !== undefined) {
      this.spellcheckerComponent.ignoredHighlights.push(this.error.errorUniqueId)
    }

    this.focusElement("toggle");
  }

  async ignorePermanently(word: string) {
    try {
      const accessTokenWithTimestamp = await this.authService.getAccessTokenWithTimestamp();
      if (!accessTokenWithTimestamp) {
        throw new Error('Valid access token not available');
      }

      await this.ignoreService.ignoreWordPermanently(
        word,
        accessTokenWithTimestamp.token
      );

      if (this.error !== undefined) {
        this.spellcheckerComponent.ignoredHighlights.push(this.error.errorUniqueId)
      }
    } catch (error) { }
    this.focusElement("toggle");
  }

  get containerStyle() {
    return {
      backgroundColor: this.getExplanationColor(this.error?.details.gravity, this.error?.details.subcategory),
      // TODO how to prevent this not being not enough? minHeight can cause jumping on mouseover
      height: "140px",
    };
  }

  getExplanationColor(gravity: number | undefined, subcategory: string | undefined): string {
    if (subcategory === 'corporate_rules') return '#A1BEED';
    if (!gravity) return "#D3E4AC";
    else if (gravity < 1.5) return "#F7D4D4";
    else if (gravity > 2.5) return "#FFFFD3";
    else return "#F8E7CB";
  }

  sendErrorToSentry() {
    Sentry.captureException(
      new Error(
        `Error word: ${this.error?.word} not found in paragraph: ${this.error?.details?.context}`
      )
    );
  }

  focusElement(id: string) {
    setTimeout(() => {
      const elementToRemove = document.getElementById(id);
      if (elementToRemove) {
        const parent = elementToRemove.parentElement;
        if (parent) {
          const siblings = Array.from(parent.children);
          const index = siblings.indexOf(elementToRemove);

          // Remove the element from the DOM
          elementToRemove.remove();

          // Find the closest suggestion element
          let elementToFocus: HTMLElement | null = null;
          if (index > 0) {
            elementToFocus = siblings[index - 1] as HTMLElement;
          } else if (index < siblings.length - 1) {
            elementToFocus = siblings[index] as HTMLElement;
          }

          // Move focus to the closest suggestion element
          if (elementToFocus) {
            elementToFocus.focus();
          }
        }
      }
    }, 100);
  }
}
