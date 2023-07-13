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

  @Input() data: any;
  // @Input() userIsSignedIn: boolean;
  @Output() showLearningBiteChange = new EventEmitter<boolean>();

  @Output()
  highlightEvent = new EventEmitter();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ suggestion: IAlternatives }>();

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

  showLearningBite = false;

  onClick() {
    this.showLearningBiteChange.emit(!this.showLearningBite);
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

  get textContainerStyle() {
    return {
      display: 'flex',
      flexDirection: this.showLearningBite ? 'row' : 'column',
      alignItems: this.showLearningBite ? 'center' : 'flex-start',
    };
  }

  get urlContainerStyle() {
    return {
      marginTop: this.showLearningBite ? '0em' : '1em',
    };
  }
}
