import { Component, EventEmitter, Input, OnDestroy, OnInit, Output, ChangeDetectionStrategy } from '@angular/core';
import { ISpellingError } from "../../data/data-structures";
import { SettingsService } from "../../services/settings.service";
import { Subscription } from "rxjs";
import { IAlternative } from "../../data/types";

@Component({
    selector: 'app-errors-list',
    templateUrl: './errors-list.component.html',
    styleUrls: ['./errors-list.component.scss'],
    changeDetection: ChangeDetectionStrategy.Eager,
    standalone: false
})
export class ErrorsListComponent implements OnInit, OnDestroy {

  @Input()
  highlights: ISpellingError[] = [];

  @Input()
  paragraphs: Map<string, string> = new Map<string, string>();

  @Output()
  highlightEvent = new EventEmitter<{ paragraphUniqueId: string, errorUniqueId: string }>();

  @Output()
  acceptSuggestionEvent = new EventEmitter<{ paragraphUniqueId: string, errorUniqueId: string, suggestion: IAlternative }>();

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

  sendHighlight(paragraphUniqueId: string, errorUniqueId: string) {
    this.highlightEvent.emit({ paragraphUniqueId, errorUniqueId });
  }

  acceptSuggestion(paragraphUniqueId: string, errorUniqueId: string, childObj: { suggestion: IAlternative }) {
    this.acceptSuggestionEvent.emit({ paragraphUniqueId, errorUniqueId, suggestion: childObj.suggestion });
  }
}
