import { ICheckResponseResult } from "./types";

export interface ISpellingError {
    paragraphUniqueId: string;
    paragraph: number;
    offset: number;
    length: number;
    word: string;
    details: ICheckResponseResult;
}
