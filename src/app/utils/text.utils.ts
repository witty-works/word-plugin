export default class TextUtils {
    static getContext(searchTerm: string, text: string): string | undefined {
        const regExpString = String.raw`((?:[a-z]+[^a-z]+)|(?:[^a-z]+[a-z]+)){0,3}(${TextUtils.escapeRegExp(searchTerm)})((?:[a-z]+[^a-z]+)|(?:[^a-z]+[a-z]+)){0,3}`;
        const regExp = new RegExp(regExpString, 'gm');

        const matches = text.match(regExp);
        if (!matches) return undefined;

        const forbiddenChars = /[äöüÄÖÜàéèòìÀÉÈÒÌ\w]/;

        for (const match of matches) {
            const index = text.indexOf(match);

            const charBefore = text[index - 1] || '';
            const charAfter = text[index + match.length] || '';

            if (!forbiddenChars.test(charBefore) && !forbiddenChars.test(charAfter)) {
                return match;
            }
        }

        return undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }
}
