import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsWithIds: { text: string, id: string }[]): string | undefined {
        const searchTerm = error.word;
        const paragraph = paragraphsWithIds.find((paragraph) => paragraph.id === error.paragraphUniqueId);
        if (!paragraph) {
            return undefined;
        }
        const text = paragraph.text;
        const escapedSearchTerm = TextUtils.escapeRegExp(searchTerm);
        const regExpString = `((?:\\p{L}+[^\\p{L}]+){0,3})\\b${escapedSearchTerm}\\b((?:[^\\p{L}]+\\p{L}+){0,3})`;
        const regExp = new RegExp(regExpString, 'gmu');

        regExp.lastIndex = Math.max(0, error.offset - 100);  // start 100 characters back, adjust as necessary

        const match = regExp.exec(text);
        if (match) {
            return match[0];
        }

        return undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
