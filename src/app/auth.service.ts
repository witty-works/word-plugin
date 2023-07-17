import { Injectable } from '@angular/core';
import { BehaviorSubject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  private _isLoggedIn = new BehaviorSubject<boolean>(this.hasToken());

  get isLoggedIn() {
    return this._isLoggedIn.asObservable(); // Returns Observable
  }

  private hasToken(): boolean {
    return !!localStorage.getItem('access_token');
  }

  login(access_token: string, refresh_token: string): void {
    localStorage.setItem('access_token', access_token);
    localStorage.setItem('refresh_token', refresh_token);
    this._isLoggedIn.next(true);
  }

  logout(): void {
    localStorage.setItem('access_token', '');
    localStorage.setItem('refresh_token', '');
    this._isLoggedIn.next(false);
  }

  getValue(): boolean {
    return this._isLoggedIn.getValue();
  }
}
