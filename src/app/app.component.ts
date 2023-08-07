import { Component } from '@angular/core';
import { TabType } from "./data/tabs";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  selectedTab: TabType = 'spellchecker';

  tabChanged(type: TabType) {
    this.selectedTab = type;
    this.isLoggedin = !!localStorage.getItem('access_token'); 
  }

  isLoggedin = true;

  ngOnInit() {
    this.isLoggedin = !!localStorage.getItem('access_token'); 
  }
}
