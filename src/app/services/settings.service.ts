import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from "rxjs";

@Injectable({
  providedIn: 'root'
})
export class SettingsService {

  private SHOW_CONTEXT_KEY = 'show-context';

  private showContext = new BehaviorSubject<boolean>(true);

  constructor() {
    const showCtx = this.loadBoolean(this.SHOW_CONTEXT_KEY);
    if (showCtx !== undefined) {
      this.showContext.next(showCtx);
    }
  }
  

  getShowContextObservable(): Observable<boolean> {
    return this.showContext.asObservable();
  }

  setShowContext(value: boolean) {
    this.showContext.next(value);
    this.save(this.SHOW_CONTEXT_KEY, value);
  }
  
  private save(name: string, value: any): void {
    localStorage.setItem(name, value);
  }

  private load(name: string): any {
    return localStorage.getItem(name);
  }

  private remove(name: string): void {
    localStorage.removeItem(name);
  }

  private loadBoolean(name: string): boolean | undefined {
    const value = this.load(name);
    if (!value) {
      return undefined;
    }
    return JSON.parse(value) === true;
  }
}
