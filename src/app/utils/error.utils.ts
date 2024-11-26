import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class ErrorUtils {
  static updateErrorMessages(lang: any, except: string | null = null) {
    const errorMessages = new Map();

    errorMessages.set("throttle-warning", "throttleWarning");
    errorMessages.set("warn-not-signed-in-word", "notSignedInWarning");
    errorMessages.set("warn-not-supported-account", "notSupportedAccountWarning");
    errorMessages.set("trial-expired-message", "trialExpired");
    errorMessages.set("issue-checking-text", "issueCheckingText");
    errorMessages.set("warn-failed-ignore-error", "failedRequestText");
    errorMessages.set("cant-identify-language", "cantIdentify");
    errorMessages.set("warn-server-error", "serverError");

    for (let [key, value] of errorMessages) {
      const errorElement = document.getElementById(key);
      if (errorElement) {
        if (except === key) {
          ErrorUtils.displayErrorMessage(errorElement, value, lang);
        } else {
          ErrorUtils.removeErrorMessage(errorElement, value, lang);
        }
      }
    }
  }

  private static createErrorMessageElement(errorMessageText: string, lang: any): HTMLElement {
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

  private static displayErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any, additionalMessage = "") {
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

  private static removeErrorMessage(messageContainer: HTMLElement, errorType: string, lang: any) {
    const errorMessageElement = messageContainer.querySelector('.error-message');
    if (errorMessageElement && errorMessageElement.textContent === lang[errorType]) {
      messageContainer.removeChild(errorMessageElement);
      if (messageContainer.children.length === 0) {
        messageContainer.style.display = 'none';
      }
    }
  }
}