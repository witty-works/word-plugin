import { environment } from 'src/environments/environment';
import { en, de, fr } from './../translations';

export function getLanguageModule(): any {
  const langCode = Office.context?.displayLanguage?.split('-')[0].toLowerCase();

  if (environment.supportedLanguages.includes(langCode)) {
    switch (langCode) {
      case 'de':
        return de;
      case 'fr':
        return fr;
      // Add other cases for supported languages as needed
    }
  }
  return en; // Fallback to English if the language code is not supported
}