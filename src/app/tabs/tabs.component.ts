import { Component, EventEmitter, Input, Output, OnInit, OnDestroy } from '@angular/core';
import { TabType } from "../data/tabs";
import { Subscription } from 'rxjs';
import { AuthService } from '../auth.service'; // import AuthService

@Component({
  selector: 'app-tabs',
  templateUrl: './tabs.component.html',
  styleUrls: ['./tabs.component.scss']
})
export class TabsComponent implements OnInit, OnDestroy {
  @Input()
  selectedTab: TabType = 'spellchecker';

  @Output()
  tabChangedEvent = new EventEmitter<TabType>();

  private isLoggedInSubscription?: Subscription;

  constructor(private authService: AuthService) {}

  tabChanged(type: TabType) {
    this.tabChangedEvent.emit(type);
  }

  isLoggedIn(): boolean {
    return this.authService.getValue();
  }

  ngOnInit() {
    this.isLoggedInSubscription = this.authService.isLoggedIn.subscribe((isLoggedIn: any) => {
      this.tabChanged(isLoggedIn ? 'spellchecker' : 'login');
    });
  }

  ngOnDestroy() {
    if (this.isLoggedInSubscription) {
      this.isLoggedInSubscription.unsubscribe();
    }
  }
}
