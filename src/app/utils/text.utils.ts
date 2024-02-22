export default class TextUtils {
    static getContext(searchTerm: string, text: string): string | undefined {
        const escapedSearchTerm = TextUtils.escapeRegExp(searchTerm);
        const regExpString = `((?:[a-z]+[^a-z]+)|(?:[^a-z]+[a-z]+)){0,3}${escapedSearchTerm}((?:[a-z]+[^a-z]+)|(?:[^a-z]+[a-z]+)){0,3}`;
        const regExp = new RegExp(regExpString, 'gm');

        let match;
        while ((match = regExp.exec(text)) !== null) {
            // Check the characters immediately before and after the match to ensure they're not alphabetic
            const matchIndex = match.index;
            const matchLength = match[0].length;
            const charBefore = (matchIndex > 0) ? text[matchIndex - 1] : '';
            const charAfter = (matchIndex + matchLength < text.length) ? text[matchIndex + matchLength] : '';
            
            const forbiddenChars = /[äöüÄÖÜàéèòìÀÉÈÒÌ]/;
            const alphaChars = /[a-z]/i; // Check for any alphabetic character, case insensitive

            if (!alphaChars.test(charBefore) && !alphaChars.test(charAfter) &&
                !forbiddenChars.test(charBefore) && !forbiddenChars.test(charAfter)) {
                return match[0]; // Return the whole match
            }
        }

        return undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
