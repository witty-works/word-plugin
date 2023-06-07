import { Component, EventEmitter, Input, Output } from '@angular/core';
import { ISpellingError } from "../../data/data-structures";
import TextUtils from "../../utils/text.utils";
import { CheckingService } from "../../services/checking.service";
import {IAlternatives} from "../../data/types";

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

  @Output()
  ignoreWordEvent = new EventEmitter<{ word: string }>();

  isOpen = false;

  suggestions: IAlternatives[] = [];


  constructor(private spellcheckerService: CheckingService) {
  }

  getContext(word: string) {
    return TextUtils.getContext(word, (this.context)!);
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

  ignoreWord(word: string) {
    this.ignoreWordEvent.emit({ word});
  }
}
