import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsWithIds: { text: string, id: string }[], allHighlights: ISpellingError[]): string | undefined {
        const searchTerm = error.word;
        // console.log('error', error);
        // console.log('paragraphsWithIds', paragraphsWithIds);
        //exclude text that is a part of another error 
        const text = paragraphsWithIds.find((paragraph) => paragraph.id === error.paragraphUniqueId)?.text;
        // console.log('text', text);
        if (!text) {
            return undefined;
        }
        const escapedSearchTerm = TextUtils.escapeRegExp(searchTerm);
        const regExpString = `((?:\\p{L}+[^\\p{L}]+)|(?:[^\\p{L}]+\\p{L}+)){0,3}\\b${escapedSearchTerm}\\b((?:\\p{L}+[^\\p{L}]+)|(?:[^\\p{L}]+\\p{L}+)){0,3}`;
        const regExp = new RegExp(regExpString, 'gmu');
        let match = regExp.exec(text);
        return match ? match[0] : undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
