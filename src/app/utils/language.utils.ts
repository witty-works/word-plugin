import { environment } from 'src/environments/environment';
import { en, de } from './../translations';

// Define a dictionary to map language codes to language modules
const LANGUAGE_MODULES: { [key: string]: any } = {
  de: de,
  en: en,
  // Add other language mappings here
};

export function getLanguageModule(): any {
  const langCode = Office.context?.displayLanguage?.split('-')[0].toLowerCase();

  if (environment.supportedLanguages.includes(langCode)) {   // Check if the language code is in the supported languages list
    return LANGUAGE_MODULES[langCode] || en; // Return the corresponding language module or fallback to English
  }
  return en;   // Fallback to English if the language code is not supported
}