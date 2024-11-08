import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsWithIds: Map<string, string>): string | undefined {
        const paragraph = paragraphsWithIds.get(error.paragraphUniqueId);
        if (!paragraph) {
            return undefined;
        }
        return paragraph.substring(error.offset - 10, error.offset + error.length + 10);
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
