import OfficePlatformType = Office.PlatformType;

class DocumentUtils {
    static async fetchParagraph(context: Word.RequestContext, paraText: string, paraPosition: number): Promise<Word.Range> {
        const docBody = context.document.body;
        context.load(docBody, 'paragraphs');
        await context.sync();

        const allParas = docBody.paragraphs.items;
        if (paraPosition < 0 || paraPosition >= allParas.length) {
            throw new Error('Incorrect paragraph position');
        }

        const desiredPara = allParas[paraPosition];
        desiredPara.load('text');
        await context.sync();

        if (!desiredPara.text.includes(paraText)) {
            throw new Error('Paragraph not present in the given context');
        }

        const paraBounds = desiredPara.getRange();
        paraBounds.load('isNullObject');
        await context.sync();

        if (paraBounds.isNullObject) {
            throw new Error('Cannot determine bounds for the paragraph');
        }

        return paraBounds;
    }

    static async fetchTextBounds(context: Word.RequestContext, withinRange: Word.Range, lookupText: string): Promise<Word.Range> {
        withinRange.load('text');
        await context.sync();

        let matchEntireWord = true;
        const lookupTextFirstCharSpecial = lookupText.charAt(0).match(/[_:*/-]/g);

        if (Office.context.diagnostics.platform === OfficePlatformType.OfficeOnline || lookupText.includes('\u000b') || lookupText.includes(' ') || lookupTextFirstCharSpecial) {
            const specialCharPattern = /\W/g;
            const match = lookupText.match(specialCharPattern) ?? lookupTextFirstCharSpecial;
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

export default DocumentUtils;