import { Component, EventEmitter, Input, OnDestroy, OnInit, Output } from '@angular/core';
import { ISpellingError } from "../../data/data-structures";
import { SettingsService } from "../../services/settings.service";
import { Subscription } from "rxjs";
import { IAlternatives } from "../../data/types";

@Component({
  selector: 'app-errors-list',
  templateUrl: './errors-list.component.html',
  styleUrls: ['./errors-list.component.scss']
})
export class ErrorsListComponent implements OnInit, OnDestroy {

  @Input()
  highlights: ISpellingError[] = [];

  @Input()
  paragraphs: { text: string, id: string }[] = [];

  @Output()
  highlightEvent = new EventEmitter<{ paragraphIndex: number, errorIndex: number }>();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ paragraphIndex: number, errorIndex: number, suggestion: IAlternatives }>();

  showContext = true;

  private settingsServiceSubscription?: Subscription;

  constructor(private settingsService: SettingsService) {
  }

  ngOnInit() {
    this.settingsServiceSubscription = this.settingsService.getShowContextObservable().subscribe(value => {
      this.showContext = value;
    });
  }

  ngOnDestroy() {
    if (this.settingsServiceSubscription) {
      this.settingsServiceSubscription.unsubscribe();
    }
  }

  sendHighlight(paragraphIndex: number, errorIndex: number) {
    this.highlightEvent.emit({ paragraphIndex, errorIndex });
  }

  acceptSuggestion(paragraphIndex: number, errorIndex: number, childObj: { suggestion: IAlternatives }) {
    this.acceptSuggestionEvent.emit({ paragraphIndex, errorIndex, suggestion: childObj.suggestion });
  }
}
