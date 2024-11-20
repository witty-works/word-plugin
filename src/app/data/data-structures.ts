import { ICheckResponseResult, IRephrasingResult } from "./types";

export interface ISpellingError {
    errorUniqueId: string;
    paragraphUniqueId: string;
    offset: number;
    length: number;
    word: string;
    details: ICheckResponseResult;
    rephrasings?: IRephrasingResult;
}