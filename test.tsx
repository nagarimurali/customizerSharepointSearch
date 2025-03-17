/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import {
    Dropdown,
    IDropdownOption,
    TextField,
    PrimaryButton,
    DetailsList,
    DetailsListLayoutMode,
    Spinner,
    MessageBar,
    MessageBarType,
    IColumn
} from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IListColumn {
    key: string;
    text: string;
    fieldType: string;
    lookupListId?: string;
    lookupField?: string;
}

interface ISearchResults {
    Title: string;
    DocType?: { Title: string };
    Status: string;
    Note: string;
    BU: string;
    [key: string]: any;
}

interface ISearchState {
    columns: IListColumn[];
    selectedColumn: string;
    query: string;
    results: ISearchResults[];
    loading: boolean;
    error: string | null;
}

interface ISearchProps {
    context: WebPartContext;
    listName: string;
}

class SearchComponent extends React.Component<ISearchProps, ISearchState> {
    constructor(props: ISearchProps) {
        super(props);
        this.state = {
            columns: [],
            selectedColumn: "",
            query: "",
            results: [],
            loading: false,
            error: null,
        };
    }

    componentDidMount() {
        this.loadColumns();
    }

    loadColumns = async () => {
        try {
            const listUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.props.listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;

            const response = await this.props.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            const data = await response.json();

            const columns = data.value.map((field: any) => ({
                key: field.InternalName,
                text: field.Title,
                fieldType: field.TypeAsString,
                lookupListId: field.LookupList,
                lookupField: field.LookupField
            }));

            this.setState({ columns });
        } catch (err) {
            this.setState({
                error: `Failed to load columns: ${err.message}`,
                columns: []
            });
        }
    };

    handleSearch = async () => {
        const { selectedColumn, query } = this.state;
        if (!selectedColumn || !query) return;

        try {
            this.setState({ loading: true, error: null, results: [] });

            const selectedColumnInfo = this.state.columns.find(col => col.key === selectedColumn);
            if (!selectedColumnInfo) throw new Error("Selected column not found");

            if (selectedColumnInfo.fieldType === 'Lookup') {
                await this.handleLookupSearch(selectedColumnInfo, query);
            } else {
                await this.handleStandardSearch(selectedColumn, query);
            }
        } catch (err) {
            this.setState({
                error: `Search failed: ${err.message}`,
                loading: false,
                results: []
            });
        }
    };

    private handleLookupSearch = async (columnInfo: IListColumn, query: string) => {
        if (!columnInfo.lookupListId) {
            throw new Error("Lookup list ID not found for this column");
        }

        const lookupField = columnInfo.lookupField || 'Title';
        const lookupListUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${columnInfo.lookupListId}')/items?$filter=${lookupField} eq '${query}'&$select=Id`;

        const lookupResponse = await this.props.context.spHttpClient.get(lookupListUrl, SPHttpClient.configurations.v1);
        const lookupData = await lookupResponse.json();

        if (!lookupData.value || lookupData.value.length === 0) {
            throw new Error("No matching items found in the lookup list");
        }

        const lookupId = lookupData.value[0].Id;

        const searchUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=Title,DocType/Title,Status,Note,BU,${columnInfo.key}/Title&$expand=DocType,${columnInfo.key}&$filter=${columnInfo.key}/Id eq ${lookupId}`;

        const response = await this.props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        this.setState({ results: data.value || [], loading: false });
    };

    private handleStandardSearch = async (columnName: string, query: string) => {
        const searchUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.props.listName)}')/items?$select=Title,DocType/Title,Status,Note,BU&$expand=DocType`;

        const response = await this.props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        const filteredResults = data.value.filter((item: any) =>
            item[columnName]?.toString().toLowerCase().includes(query.toLowerCase())
        );

        this.setState({ results: filteredResults, loading: false });
    };

    handleColumnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        this.setState({ selectedColumn: option?.key.toString() || "" });
    };

    handleQueryChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        this.setState({ query: newValue || "" });
    };

    render() {
        const { columns, selectedColumn, query, results, loading, error } = this.state;

        const columnsConfig: IColumn[] = [
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
                key: 'Note',
                name: 'Note',
                fieldName: 'Note',
                minWidth: 150,
                isResizable: true
            },
            {
                key: 'BU',
                name: 'BU',
                fieldName: 'BU',
                minWidth: 100,
                isResizable: true
            }
        ];

        return (
            <div className="search-container" style={{ padding: 20 }}>
                <Dropdown
                    label="Select search column"
                    options={columns.map(c => ({ key: c.key, text: c.text }))}
                    selectedKey={selectedColumn}
                    onChange={this.handleColumnChange}
                    styles={{ dropdown: { width: 300 } }}
                />

                <TextField
                    label="Search value"
                    value={query}
                    onChange={this.handleQueryChange}
                    disabled={!selectedColumn}
                    styles={{ root: { marginTop: 15 } }}
                />

                <PrimaryButton
                    text="Search"
                    onClick={this.handleSearch}
                    disabled={!selectedColumn || !query}
                    styles={{ root: { marginTop: 15, marginBottom: 20 } }}
                />

                {loading && <Spinner label="Searching..." />}

                {error && (
                    <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 15 } }}>
                        {error}
                    </MessageBar>
                )}

                {results.length > 0 && (
                    <DetailsList
                        items={results}
                        columns={columnsConfig}
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        styles={{ root: { marginTop: 20 } }}
                    />
                )}

                {!loading && !error && results.length === 0 && (
                    <MessageBar styles={{ root: { marginTop: 15 } }}>
                        No results found
                    </MessageBar>
                )}
            </div>
        );
    }
}

export default SearchComponent;
