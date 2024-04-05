import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsWithIds: { text: string, id: string }[]): string | undefined {
        const searchTerm = error.word;
        const text = paragraphsWithIds.find((paragraph) => paragraph.id === error.paragraphUniqueId)?.text;
        if (!text) {
            return undefined;
        }
        const escapedSearchTerm = TextUtils.escapeRegExp(searchTerm);
        const regExpString = `((?:\\p{L}+[^\\p{L}]+)|(?:[^\\p{L}]+\\p{L}+)){0,3}${escapedSearchTerm}((?:\\p{L}+[^\\p{L}]+)|(?:[^\\p{L}]+\\p{L}+)){0,3}`;
        const regExp = new RegExp(regExpString, 'gmu');
        let match = regExp.exec(text)
        return match ? match[0] : undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
