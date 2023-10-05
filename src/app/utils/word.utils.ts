import OfficePlatformType = Office.PlatformType;

import PlatformType = Office.PlatformType;

export default class DocumentUtils {
    static async fetchParagraph(context: Word.RequestContext, paragraph: string): Promise<Word.Range> {
        const body = context.document.body;
        context.load(body);
        await context.sync();

        const chunks = DocumentUtils.divideTextIntoSegments(paragraph, 255);

        let fullRange: Word.Range | null = null;
        for (let index = 0; index < chunks.length; index++) {
            const chunk = chunks[index];

            if (chunk.trim() === "") {
                continue;
            }

            const paragraphRangeCollection = body.search(chunk, {
                matchCase: true,
            });

            const paragraphRange = paragraphRangeCollection.getFirstOrNullObject(); //get paragraph at index here        

            paragraphRange.load('isNullObject');

            await context.sync();

            if (!paragraphRange || paragraphRange.isNullObject) {
                return Promise.reject(new Error('Could not find range for chunk: ' + chunk));

            }

            if (!fullRange) {
                fullRange = paragraphRange;
            } else {
                fullRange = fullRange.expandTo(paragraphRange);
            }

        }

        if (!fullRange) {
            return Promise.reject(new Error('Context paragraph not found'));
        }

        return fullRange;
    }    

    static async fetchTextBounds(context: Word.RequestContext, withinRange: Word.Range, lookupText: string): Promise<Word.Range> {
        withinRange.load('text');
        await context.sync();

        let matchEntireWord = true;
        if (Office.context.diagnostics.platform === OfficePlatformType.OfficeOnline || lookupText.includes('\u000b') || lookupText.includes(' ')) {
            const specialCharPattern = /\W/g;
            const match = lookupText.match(specialCharPattern);
            if (match !== null) {
                matchEntireWord = false;
                console.log(`Whole word matching disabled for: ${lookupText}`);
            }
        }

        const searchTextRanges = withinRange.search(lookupText, {
            matchCase: true,
            matchWholeWord: matchEntireWord,
        });

        const locatedTextRange = searchTextRanges.getFirstOrNullObject();
        locatedTextRange.load('isNullObject');
        await context.sync();

        if (!locatedTextRange || locatedTextRange.isNullObject) {
            throw new Error(`Cannot find the range for text: ${lookupText}`);
        }

        return locatedTextRange;
    }

    static divideTextIntoSegments(text: string, segmentSize: number): string[] {
        const segments: string[] = [];

        let segmentBuilder: string = '';
        let count: number = 1;
        for (const character of text) {
            const isInvalid = DocumentUtils.checkInvalidCharacter(character);
            if (count > segmentSize || (count > 0 && isInvalid)) {
                segments.push(segmentBuilder);
                segmentBuilder = '';
                count = 1;
            }
            if (!isInvalid) {
                segmentBuilder += character;
                count++;
            }
        }

        segments.push(segmentBuilder);

        return segments;
    }

    static checkInvalidCharacter(character: string): boolean {
        const charCode = character.charCodeAt(0);
        return (charCode >= 0 && charCode <= 0x1F) || charCode === 0x7f || (charCode >= 0x80 && charCode <= 0x9F);
    }
}