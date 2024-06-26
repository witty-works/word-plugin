import { Injectable } from '@angular/core';
import { de, en } from '../translations';

@Injectable({
  providedIn: 'root'
})
export class ErrorUtils {
  static createErrorMessageElement(errorMessageText: string): HTMLElement {
  
    const lang = Office.context?.displayLanguage?.split('-')[0].toLowerCase() === 'de' ? de : en;

    const message = document.createElement('div');
    const cautionImg = document.createElement('img');
    cautionImg.src = "assets/icons/caution-sign.svg";
    cautionImg.alt = lang.cautionIconAlt;
    cautionImg.style.marginRight = '5px';
    message.appendChild(cautionImg);
    
    const errorMessageTextNode = document.createTextNode(errorMessageText);
    message.appendChild(errorMessageTextNode);

    return message;
  }
}