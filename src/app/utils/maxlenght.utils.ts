import { Injectable } from '@angular/core';
import { de, en } from '../translations';

@Injectable({
  providedIn: 'root'
})
export class MaxLengthUtils {
  static createErrorMessageElement(errorMessageText: string): HTMLElement {
    // Determine the language based on the Office context
    const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

    // Create the main message container
    const message = document.createElement('div');
    message.classList.add('max-length-reached-banner');
    message.setAttribute('role', 'alert');
    message.setAttribute('aria-live', 'assertive');

    // Create and append the caution image
    const cautionImg = document.createElement('img');
    cautionImg.src = "assets/icons/caution-sign.svg";
    cautionImg.alt = lang.cautionIconAlt;
    cautionImg.style.marginRight = '5px';
    message.appendChild(cautionImg);

    // Create and append the link that triggers highlighting
    const link = document.createElement('a');
    link.href = "#";
    link.className = "max-length-reached-banner-link";
    link.setAttribute('role', 'link');
    link.tabIndex = 0;
    link.setAttribute('aria-label', lang.hitMaxTextLengthHere);
    link.setAttribute('aria-describedby', "maxTextLengthDescription");
    link.innerHTML = lang.hitMaxTextLengthHere;
    // link.addEventListener('click', (event) => {
    //   event.preventDefault();
    //   highlightCheckedText(); // This function needs to be defined elsewhere
    // });
    // link.addEventListener('keyup', (event) => {
    //   if (event.key === 'Enter') {
    //     highlightCheckedText(); // This function needs to be defined elsewhere
    //   }
    // });
    message.appendChild(link);

    // Create and append the description span
    const descriptionSpan = document.createElement('span');
    descriptionSpan.id = "maxTextLengthDescription";
    descriptionSpan.textContent = lang.hitMaxTextLength2;
    message.appendChild(descriptionSpan);

    // Append the error message text node
    const errorMessageTextNode = document.createTextNode(errorMessageText);
    message.appendChild(errorMessageTextNode);

    return message;
  }

static displayErrorMessage(messageContainer: HTMLElement, errorTypes: string[], lang: any) {
  // Check for an existing error message container
  let errorMessageContainer = messageContainer.querySelector('.error-messages-container');
  if (!errorMessageContainer) {
    errorMessageContainer = document.createElement('div');
    errorMessageContainer.className = 'error-messages-container';
    messageContainer.appendChild(errorMessageContainer);
  } else {
    // Clear existing messages
    errorMessageContainer.innerHTML = '';
  }

  // Display the message container
  messageContainer.style.display = 'block';

  // Iterate over each error type and create an error message element
  errorTypes.forEach(errorType => {
    const errorMessageElement = this.createErrorMessageElement(lang[errorType]);
    if (errorMessageContainer) {
      errorMessageContainer.appendChild(errorMessageElement);
    } else {
      console.error("errorMessageContainer is null");
    }
  });

  // Add an event listener to remove the error message when the user comes back online
  window.addEventListener('online', () => {
    if (errorMessageContainer && messageContainer.contains(errorMessageContainer)) {
      messageContainer.removeChild(errorMessageContainer);
      messageContainer.style.display = 'none';
    }
  });
}
}