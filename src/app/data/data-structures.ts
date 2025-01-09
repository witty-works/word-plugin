import { ICheckResponseResult, IRephrasingResult } from "./types";

export interface ISpellingError {
    errorUniqueId: string;
    paragraphUniqueId: string;
    offset: number;
    length: number;
    word: string;
    gender_separator: string;
    details: ICheckResponseResult;
    rephrasings?: IRephrasingResult;
}