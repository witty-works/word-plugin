import { ICheckResponseResult } from "./types";

export interface ISpellingError {
    paragraph: number;
    offset: number;
    length: number;
    word: string;
    details: ICheckResponseResult;
}
