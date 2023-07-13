import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from "rxjs";

@Injectable({
  providedIn: 'root'
})
export class SettingsService {

  private SHOW_CONTEXT_KEY = 'show-context';

  private showContext = new BehaviorSubject<boolean>(true);
  private checkUpperAndLowerCase = new BehaviorSubject<boolean>(true);
  private checkGrammarAndSpelling = new BehaviorSubject<boolean>(true);

  constructor() {
    const showCtx = this.loadBoolean(this.SHOW_CONTEXT_KEY);
    if (showCtx !== undefined) {
      this.showContext.next(showCtx);
    }

    const checkUpperAndLowerCase = this.loadBoolean('check-upper-and-lower-case');
    if (checkUpperAndLowerCase !== undefined) {
      this.checkUpperAndLowerCase.next(checkUpperAndLowerCase);
    }

    const checkGrammarAndSpelling = this.loadBoolean('check-grammar-and-spelling');
    if (checkGrammarAndSpelling !== undefined) {
      this.checkGrammarAndSpelling.next(checkGrammarAndSpelling);
    }
  }
  

  getShowContextObservable(): Observable<boolean> {
    return this.showContext.asObservable();
  }

  getCheckUpperAndLowerCaseObservable(): Observable<boolean> {
    return this.checkUpperAndLowerCase.asObservable();
  }

  getCheckGrammarAndSpellingObservable(): Observable<boolean> {
    return this.checkGrammarAndSpelling.asObservable();
  }

  setShowContext(value: boolean) {
    this.showContext.next(value);
    this.save(this.SHOW_CONTEXT_KEY, value);
  }

  setCheckUpperAndLowerCase(value: boolean) {
    this.checkUpperAndLowerCase.next(value);
    this.save('check-upper-and-lower-case', value);
  }

  setCheckGrammarAndSpelling(value: boolean) {
    this.checkGrammarAndSpelling.next(value);
    this.save('check-grammar-and-spelling', value);
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
