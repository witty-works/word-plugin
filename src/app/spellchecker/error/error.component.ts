import { Component, EventEmitter, Input, Output, HostListener, WritableSignal, signal } from "@angular/core";
import { ISpellingError, } from "../../data/data-structures";
import TextUtils from "../../utils/text.utils";
import { CheckingService } from "../../services/checking.service";
import { IgnoreService } from "../../services/ignore.service";
import { AuthService } from '../../services/auth.service';
import { SpellcheckerComponent } from "../spellchecker.component";
import { IAlert, IAlternative, IRephrasingResult } from "../../data/types";
import { useAnalytics } from "src/app/analytics/analytics";
import * as Sentry from "@sentry/browser";
import { KEYBOARD_SHORTCUTS_CONFIG } from "src/app/keyboard-shortcuts.config";
import { getLanguageModule } from '../../utils/language.utils';

const analytics = useAnalytics();

@Component({
  selector: "app-error",
  templateUrl: "./error.component.html",
  styleUrls: ["./error.component.scss"],
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

  diff = require('diff');

  rephrasings: WritableSignal<IRephrasingResult | null> = signal(null);

  constructor(
    private spellcheckerService: CheckingService,
    private spellcheckerComponent: SpellcheckerComponent,
    private authService: AuthService,
    private ignoreService: IgnoreService
  ) {
    this.lang = getLanguageModule();
  }

  getRephrasing(suggestion: IAlternative) {
    if (!this.error) {
      return false;
    }

    if (!this.spellcheckerComponent.isLLMAlternativesActive(this.error)) {
      return false;
    }

    const rephrasings: IRephrasingResult | null = this.rephrasings()
    if (rephrasings === null) {
      return 'Spinner ..';
    }

    const rephrasing = rephrasings.results.get(suggestion.text);
    if (rephrasing === undefined) {
      return false;
    }

    let diff: string;

    const diffElements: [{ added: boolean, value: string }] = this.diff.diffWords(
      rephrasings.sentence,
      rephrasing
    );

    diff = ""
    diffElements.forEach((diffElement) => {
      if (diffElement.added) {
        diff += " <b>" + diffElement.value + "</b>"
      } else if (!diff.endsWith(" ..")) {
        diff += " .."
      }
    });

    return diff;
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

    if (!this.isOpen
      || this.error === undefined // TODO handle error
      || !this.spellcheckerComponent.isLLMAlternativesActive(this.error)
    ) {
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
    };
  }

  getExplanationColor(gravity: number | undefined, subcategory: string | undefined): string {
    if (subcategory === 'corporate_rules') return '#A1BEED';
    if (!gravity) return "#D3E4AC";
    else if (gravity < 1.5) return "#F7D4D4";
    else if (gravity > 2.5) return "#FFFFD3";
    else return "#F8E7CB";
  }

  sendErrorToSentry(error?: any) {
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
