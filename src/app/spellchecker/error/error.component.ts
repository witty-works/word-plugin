import { Component, EventEmitter, Input, Output } from '@angular/core';
import { ISpellingError } from "../../data/data-structures";
import TextUtils from "../../utils/text.utils";
import { CheckingService } from "../../services/checking.service";
import {IAlternatives} from "../../data/types";
import { en, de } from '../../translations';

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

  @Input() data: any;

  @Output()
  highlightEvent = new EventEmitter();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ suggestion: IAlternatives }>();

  isOpen = false;
  suggestions: IAlternatives[] = [];
  showLearningBite: boolean = false;
  lang = Office.context.displayLanguage.split('-')[0].toLowerCase() === 'de' ? de : en;

  constructor(private spellcheckerService: CheckingService) {
  }
  
  getContext(word: string) {
    let ctxt = TextUtils.getContext(word, (this.context)!);
    if (ctxt) {
      ctxt = ctxt.replace(/()/g, '<img src="assets/icons/soft-return.svg" class="soft-return-icon" alt="Soft return icon"><br>');
    }
    return ctxt;
  }

  async toggle(): Promise<void> {
    if (!this.isOpen) {
      this.isOpen = true;
      this.suggestions = await this.spellcheckerService.getSuggestions(this.error!);
      this.sendHighlight()
    } else {
      this.isOpen = false;
    }
  }

  sendHighlight() {
    this.highlightEvent.emit();
  }

  acceptSuggestion(suggestion: IAlternatives) {
    this.acceptSuggestionEvent.emit({ suggestion });
  }


  onClick(url: string | undefined) {
    this.showLearningBite = !this.showLearningBite;
    url && window.open(url.split('?')[0], '_blank');
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
}
