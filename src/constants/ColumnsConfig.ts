import { IColumn } from "@fluentui/react";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


export const columnsConfig: IColumn[] = [
    {
        key: 'Title',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 100,
        isResizable: true
    },
    {
        key: 'DocType',
        name: 'DocType',
        minWidth: 100,
        isResizable: true,
        onRender: (item: ISearchResults) => item.DocType?.Title || "N/A"
    },
    {
        key: 'Status',
        name: 'Status',
        fieldName: 'Status',
        minWidth: 100,
        isResizable: true
    },
    {
        key: 'BU',
        name: 'BU',
        fieldName: 'BU',
        minWidth: 100,
        isResizable: true
    },
    {
        key: 'PartNumber',
        name: 'PartNumber',
        fieldName: 'PartNumber',
        minWidth: 100,
        isResizable: true
    },
    {
        key: "approvalDetails",
        name: "Approval Details",
        minWidth: 120,
        isResizable: false,
      
    }

];
