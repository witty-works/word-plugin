import { Component, EventEmitter, Input, Output, OnInit } from '@angular/core';
import { TabType } from "../data/tabs";
import { getLanguageModule } from '../utils/language.utils';

@Component({
    selector: 'app-tabs',
    templateUrl: './tabs.component.html',
    styleUrls: ['./tabs.component.scss'],
    standalone: false
})
export class TabsComponent implements OnInit {
  lang: any;

  @Input()
  selectedTab: TabType = 'spellchecker';

  @Output()
  tabChangedEvent = new EventEmitter<TabType>();

  ngOnInit() {
    this.lang = getLanguageModule();
  }

  tabChanged(type: TabType) {
    this.tabChangedEvent.emit(type);
  }
}