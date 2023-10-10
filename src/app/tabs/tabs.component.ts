import { Component, EventEmitter, Input, Output } from '@angular/core';
import { TabType } from "../data/tabs";
import { en, de } from '../translations';

@Component({
  selector: 'app-tabs',
  templateUrl: './tabs.component.html',
  styleUrls: ['./tabs.component.scss']
})
export class TabsComponent {
  lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

  @Input()
  selectedTab: TabType = 'spellchecker';

  @Output()
  tabChangedEvent = new EventEmitter<TabType>();

  tabChanged(type: TabType) {
    this.tabChangedEvent.emit(type);
  }
}
