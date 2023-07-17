import { Component, EventEmitter, Input, Output } from '@angular/core';
import { TabType } from "../data/tabs";
import { AuthService } from '../auth.service'; // import AuthService

@Component({
  selector: 'app-tabs',
  templateUrl: './tabs.component.html',
  styleUrls: ['./tabs.component.scss']
})
export class TabsComponent {
  @Input()
  selectedTab: TabType = 'spellchecker';

  @Output()
  tabChangedEvent = new EventEmitter<TabType>();

  constructor(private authService: AuthService) {}

  tabChanged(type: TabType) {
    this.tabChangedEvent.emit(type);
  }

  isLoggedIn(): boolean {
    return this.authService.getValue();
  }
}
