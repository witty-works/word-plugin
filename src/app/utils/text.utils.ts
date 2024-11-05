import { ISpellingError } from "../data/data-structures";

export default class TextUtils {
    static getContext(error: ISpellingError, paragraphsWithIds: { text: string, id: string }[]): string[] | undefined {
        const searchTerm = error.word;
        const paragraph = paragraphsWithIds.find((paragraph) => paragraph.id === error.paragraphUniqueId);
        if (!paragraph) {
            return undefined;
        }

        const text = paragraph.text;
        const contexts: string[] = [];

        // Split the paragraph into words, including spaces as delimiters
        const words = text.split(/\s+/);

        // Loop through the words to find all instances of the search term
        for (let i = 0; i < words.length; i++) {
            if (words[i] === searchTerm) {
                // Capture the 3 words before and after the search term
                const contextStart = Math.max(0, i - 3);
                const contextEnd = Math.min(words.length, i + 4);

                // Join the selected words into a context string
                const context = words.slice(contextStart, contextEnd).join(" ");
                contexts.push(context);
            }
        }

        // Return an array of all found contexts, or undefined if none were found
        return contexts.length > 0 ? contexts : undefined;
    }

    static escapeRegExp(strg: string) {
        return strg.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
}

// Function to search and highlight words in the paragraph

function highlightWordsInParagraph(paragraph: string, check_result: any[]) {
    let check_result_left_to_search = check_result.map(item => item.text); // Create a flat list of highlighted words
    const check_result_keys_to_hide: number[] = [];
    let count = 0;

    for (let i = 0; i < paragraph.length; i++) {
        let word = "";

        // Build the next word until a whitespace is encountered
        while (i < paragraph.length && !/\s/.test(paragraph[i])) {
            word += paragraph[i];
            i++;
        }

        // Find the index of the word in the list of words to search
        const index = check_result_left_to_search.indexOf(word);
        if (index > -1) {
            // Update start/end in check_result
            const startIdx = i - word.length;
            check_result[index].start = startIdx;
            check_result[index].end = i;

            // Only hide results if index > 0, ie. in case we need to skip items in the check_results.
            if (index > 0) {
                // Track the index of the word found for hiding
                check_result_keys_to_hide.push(count);

                // Remove words from the list that no longer need to be looked for
                let remainingPopCount = index + 1;
                while (remainingPopCount > 0 && check_result_left_to_search.length > 0) {
                    check_result_left_to_search.pop();
                    remainingPopCount--;
                }
            }
        }
        count++;

        // Skip to the next non-whitespace character
        while (i < paragraph.length && /\s/.test(paragraph[i])) {
            i++;
        }

        // End the loop early if all words are processed
        if (count === check_result.length) {
            break;
        }
    }

    return {
        updated_check_result: check_result,
        check_result_keys_to_hide: check_result_keys_to_hide,
    };
}
