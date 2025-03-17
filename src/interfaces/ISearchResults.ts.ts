/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISearchResults {
    Title: string;
    DocType?: { Title: string };
    Status: string;
    Note: string;
    BU: string;
    PartNumber: string;
    [key: string]: any;
}
