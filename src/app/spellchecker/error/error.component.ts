import { Component, EventEmitter, Input, Output } from '@angular/core';
import { ISpellingError } from "../../data/data-structures";
import TextUtils from "../../utils/text.utils";
import { CheckingService } from "../../services/checking.service";
import { SpellcheckerComponent } from "../spellchecker.component";
import {IAlert, IAlternatives} from "../../data/types";
import { en, de } from '../../translations';
import { useAnalytics } from 'src/app/analytics/analytics';
import * as Sentry from '@sentry/browser';

const analytics = useAnalytics();

@Component({
  selector: 'app-error',
  templateUrl: './error.component.html',
  styleUrls: ['./error.component.scss']
})
export class ErrorComponent {

  @Input()
  error?: ISpellingError;

  @Input()
  context?: string;

  @Input()
  showContext: boolean = true;

  @Output()
  highlightEvent = new EventEmitter();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ suggestion: IAlternatives }>();

  isOpen = false;
  suggestions: IAlternatives[] = [];
  showLearningBite: boolean = false;
  
  alerts: IAlert[] = this.spellcheckerComponent.alerts;

  paragraphsWithIds: { text: string, id: string }[] = this.spellcheckerComponent.paragraphsWithIds;
  
  lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

  constructor(private spellcheckerService: CheckingService, private spellcheckerComponent: SpellcheckerComponent) {
  }

  getContextErrorComponent(error: ISpellingError) {
    let ctxt = TextUtils.getContext(error, this.paragraphsWithIds);
    if (ctxt) {
      ctxt = ctxt.replace(/()/g, '<img src="assets/icons/soft-return.svg" class="soft-return-icon" alt="Soft return icon"><br>');
    } else {
      this.sendErrorToSentry();
    }
    return ctxt;
  }

  async toggle(): Promise<void> {
    this.suggestions = await this.spellcheckerService.getSuggestions(this.error!);
    const alertRelevantToSuggestion = this.alerts.find((alert) => {
      return alert.data.text === this.error?.word;
    });

    if (!this.isOpen) {
      this.isOpen = true;
      alertRelevantToSuggestion && analytics.popoverLogs(alertRelevantToSuggestion, 'popover_open');
      this.sendHighlight()
    } else {
      this.isOpen = false;
      alertRelevantToSuggestion && analytics.popoverLogs(alertRelevantToSuggestion, 'popover_close');
    }
  }

  sendHighlight() {
    this.highlightEvent.emit();
  }

  acceptSuggestion(suggestion: IAlternatives) {
    if (suggestion.remove) {
      suggestion.text = '';
    }
    this.acceptSuggestionEvent.emit({ suggestion });
  }


  onClick(url: string | undefined) {
    this.showLearningBite = !this.showLearningBite;
    const alertRelevantToSuggestion = this.alerts.find((alert) => {
      return alert.data.text === this.error?.word;
    });
    alertRelevantToSuggestion && analytics.popoverLogs(alertRelevantToSuggestion, 'learning_bites');
    url && Office.context.ui?.openBrowserWindow(url);
  }

  ignore() {
    this.spellcheckerComponent.highlights = this.spellcheckerComponent.highlights.filter((error) => {
      return error.word !== this.error?.word;
    });
    this.spellcheckerComponent.updatehighlights();
  }

  get containerStyle() {
    return {
      backgroundColor: this.getExplanationColor(this.error?.details.gravity),
    };
  }

  getExplanationColor(
    gravity: number | undefined,
  ): string {
    if (!gravity) return '#D3E4AC';
    else if (gravity < 1.5) return '#F7D4D4';
    else if (gravity > 2.5) return '#FFFFD3';
    else return '#F8E7CB';
  }

  sendErrorToSentry() {
    Sentry.captureException(new Error(`Error: ${this.error} not found in paragraph`));
  }
}
