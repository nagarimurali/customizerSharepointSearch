/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


export class SearchService {
    private context: WebPartContext;
    private listName: string;

    constructor(context: WebPartContext, listName: string) {
        this.context = context;
        this.listName = listName;
    }

    async loadColumns(): Promise<IListColumn[]> {
        try {
            const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
            const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            const data = await response.json();

            return data.value.map((field: any) => ({
                key: field.InternalName,
                text: field.Title,
                fieldType: field.TypeAsString,
                lookupListId: field.LookupList,
                lookupField: field.LookupField
            }));
        } catch (err) {
            throw new Error(`Failed to load columns: ${err.message}`);
        }
    }
//Update
   
    async handleLookupSearch(columnInfo: IListColumn, query: string): Promise<ISearchResults[]> {

        
        // if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found for this column");

        // const lookupField = columnInfo.lookupField || 'Title';
        // const lookupListUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=startswith(${lookupField}, '${query}')&$select=Id`;

        // const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        // const lookupData = await lookupResponse.json();

        if (!columnInfo.lookupListId) throw new Error("Lookup list ID not found");

        const lookupField = columnInfo.lookupField || 'Title';
        const lookupListUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=substringof('${query}', ${lookupField})`; // Use substringof for partial matches

        const lookupResponse = await this.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        const lookupData = await lookupResponse.json();

        if (!lookupData.value || lookupData.value.length === 0) {
            throw new Error("No matching items found in the lookup list");
        }

        const lookupIds = lookupData.value.map((item: any) => item.Id).join(",");
        const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType&$filter=${columnInfo.key}/Id in (${lookupIds})`;

        const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        return data.value || [];
    }

    // async handleStandardSearch(columnName: string, query: string): Promise<ISearchResults[]> {
    //     const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType`;

    //     const response = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
    //     const data = await response.json();

    //     return data.value.filter((item: any) =>
    //         item[columnName]?.toString().toLowerCase().includes(query.toLowerCase())
    //     );
    // }

    
    // async handleStandardSearch(filters: { columnName: string; query: string }[]): Promise<ISearchResults[]> {
    //     try {
    //         const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType`;
    //         const response: SPHttpClientResponse = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
    //         const data = await response.json();

    //         // Apply AND logic: All filters must match
    //         const filteredResults = data.value.filter((item: any) =>
    //             filters.every(filter => {
    //                 const fieldValue = item[filter.columnName]?.toString().toLowerCase() || "";
    //                 const queryValue = filter.query.toLowerCase();
    //                 return fieldValue === queryValue; // Exact match
    //                 // return fieldValue.includes(queryValue); // For partial matches
    //             })
    //         );

    //         return filteredResults;
    //     } catch (error) {
    //         throw new Error(`Search failed: ${error.message}`);
    //     }
    // }
//Update
    async handleStandardSearch(filters: { columnName: string; query: string }[]): Promise<ISearchResults[]> {
        try {
            const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.listName)}')/items?$select=Title,DocType/Title,Status,BU,PartNumber&$expand=DocType`;
            const response: SPHttpClientResponse = await this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
            const data = await response.json();

            // Case-insensitive partial matching (e.g., "bu" matches "BuTest", "TESTBU")
            const filteredResults = data.value.filter((item: any) =>
                filters.every(filter => {
                    const fieldValue = item[filter.columnName]?.toString().toLowerCase() || "";
                    const queryValue = filter.query.toLowerCase();
                    return fieldValue.includes(queryValue); // Partial match
                })
            );

            return filteredResults;
        } catch (error) {
            throw new Error(`Search failed: ${error.message}`);
        }
    }




}
