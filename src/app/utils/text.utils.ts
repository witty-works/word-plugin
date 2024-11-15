import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsByUniqueId: Map<string, string>): string | undefined {
        const paragraph = paragraphsByUniqueId.get(error.paragraphUniqueId);
        if (!paragraph) {
            return undefined;
        }
        // Get the paragraph from the last space before the error to the next whitespace after the error, so full words are shown
        const start = Math.max(0, paragraph.lastIndexOf(" ", error.offset - 10) + 1);
        const end = paragraph.indexOf(" ", error.offset + error.length + 10);
        const adjustedEnd = end === -1 ? paragraph.length : end;

        return paragraph.substring(start, adjustedEnd);
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
