// this code is based on https://github.com/divvun/divvun-gramcheck-web/blob/master/msoffice/src/utils/index.ts

import PlatformType = Office.PlatformType;

export default class WordUtils {
    static async getParagraphRange(context: Word.RequestContext, paragraph: string, paragraphIndex: number): Promise<Word.Range> {
            const body = context.document.body;
            context.load(body, 'paragraphs');
            await context.sync();
        
            const allParagraphs = body.paragraphs.items;
            if (paragraphIndex < 0 || paragraphIndex >= allParagraphs.length) {
                return Promise.reject(new Error('Invalid paragraph index'));
            }
        
            const targetParagraph = allParagraphs[paragraphIndex];
            targetParagraph.load('text');
            await context.sync();
        
            if (!targetParagraph.text.includes(paragraph)) {
                return Promise.reject(new Error('Could not find paragraph in context'));
            }
        
            const paragraphRange = targetParagraph.getRange();
            paragraphRange.load('isNullObject');
            await context.sync();
        
            if (paragraphRange.isNullObject) {
                return Promise.reject(new Error('Could not get range for paragraph'));
            }
        
            return paragraphRange;
        }


    static async getWordRange(context: Word.RequestContext, range: Word.Range, errorText: string): Promise<Word.Range> {
        range.load('text');
        range.insertText(' ', 'End');
        await context.sync();

        // Word online seems to have issues with searching whole words including special characters:
        // https://github.com/OfficeDev/office-js/issues/3360
        // Thus, disabling whole word match if online and if the word contains a special character
        let useMatchWholeWord = true;
        if (Office.context.diagnostics.platform === PlatformType.OfficeOnline || errorText.includes('\u000b')) {
            let pattern = /\W/g;
            let result = errorText.match(pattern);
            if(result !== null){
                useMatchWholeWord = false;
                console.log("Disable whole word match for word: " + errorText);
            }
        }

        const errorTextRangeCollection = range.search(errorText, {
            matchCase: true,
            matchWholeWord: useMatchWholeWord,
        });

        const foundErrorRange = errorTextRangeCollection.getFirstOrNullObject();
        foundErrorRange.load('isNullObject');
        await context.sync();

        if (!foundErrorRange || foundErrorRange.isNullObject) {
            return Promise.reject(new Error('The range for the error was not found: ' + errorText));
        }

        return foundErrorRange;
    }

    static splitStringToChunks(string: string, chunkLength: number): string[] {
        const chunks: string[] = [];

        let tempString: string = '';
        let counter: number = 1;
        for (const char of string) {
            const invalidChar = WordUtils.isInvalidSearchCharacter(char);
            if (counter > chunkLength || (counter > 0 && invalidChar)) {
                chunks.push(tempString);
                tempString = '';
                counter = 1;
            }
            if (!invalidChar) {
                tempString += char;
                counter++;
            }
        }

        chunks.push(tempString);

        return chunks;
    }

    static isInvalidSearchCharacter(char: string): boolean {
        const code = char.charCodeAt(0);
        return (code >= 0 && code <= 0x1F) || code === 0x7f || (code >= 0x80 && code <= 0x9F);
    }
}
