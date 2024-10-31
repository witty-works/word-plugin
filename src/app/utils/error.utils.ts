import { Injectable } from '@angular/core';
import { getLanguageModule } from './language.utils';

@Injectable({
  providedIn: 'root'
})
export class ErrorUtils {
  private static defaultLang = getLanguageModule();

  static createErrorMessageElement(errorMessageText: string, lang: any = ErrorUtils.defaultLang): HTMLElement {
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

  static displayErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any = ErrorUtils.defaultLang, additionalMessage = "") {
    const existingErrorMessage = messageContainer.querySelector('.error-message');
    if (additionalMessage) {
      additionalMessage = ` - ["${additionalMessage}"]`;
    }
    const errorMessageElement = this.createErrorMessageElement(lang[errorType] + additionalMessage, lang);

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

  static removeErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any = ErrorUtils.defaultLang) {
    const errorMessageElement = messageContainer.querySelector('.error-message');
    if (errorMessageElement && errorMessageElement.textContent === lang[errorType]) {
      messageContainer.removeChild(errorMessageElement);
      if (messageContainer.children.length === 0) {
        messageContainer.style.display = 'none';
      }
    }
  }
}