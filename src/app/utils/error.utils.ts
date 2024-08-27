import { Injectable } from '@angular/core';
import { de, en } from '../translations';

@Injectable({
  providedIn: 'root'
})
export class ErrorUtils {
  static createErrorMessageElement(errorMessageText: string): HTMLElement {
  
    const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

    const message = document.createElement('div');
    message.classList.add('error-message');
    message.setAttribute('role', 'alert');
    message.setAttribute('aria-live', 'assertive');
    
    const cautionImg = document.createElement('img');
    cautionImg.src = "assets/icons/caution-sign.svg";
    cautionImg.alt = lang.cautionIconAlt;
    cautionImg.style.marginRight = '5px';
    message.appendChild(cautionImg);
    
    const errorMessageTextNode = document.createTextNode(errorMessageText);
    message.appendChild(errorMessageTextNode);

    return message;
  }

  static displayErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any) {
    const existingErrorMessage = messageContainer.querySelector('.error-message');
    const errorMessageElement = this.createErrorMessageElement(lang[errorType]);

    messageContainer.style.display = 'block';
    if (existingErrorMessage) {
      messageContainer.replaceChild(errorMessageElement, existingErrorMessage);
    } else {
      messageContainer.appendChild(errorMessageElement);
    }

    window.addEventListener('online', () => {
      if (existingErrorMessage) {
        messageContainer.removeChild(existingErrorMessage);
      }
      messageContainer.style.display = 'none';
    });
  }

  static removeErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any) {
    const errorMessageElement = messageContainer.querySelector('.error-message');
    if (errorMessageElement && errorMessageElement.textContent === lang[errorType]) {
      messageContainer.removeChild(errorMessageElement);
      if (messageContainer.children.length === 0) {
        messageContainer.style.display = 'none';
      }
    }
  }
}