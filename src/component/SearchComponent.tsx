/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/no-new-null */
import * as React from "react";
import {
    Dropdown,
    IDropdownOption,
    TextField,
    PrimaryButton,
    DetailsList,
    Spinner,
    MessageBar,
    MessageBarType
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SearchService } from "../service/SearchService";
import { columnsConfig } from "../constants/ColumnsConfig";
import { IListColumn } from "../interfaces/IListColumn";
import { ISearchResults } from "../interfaces/ISearchResults.ts";


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
    private searchService: SearchService;

    constructor(props: ISearchProps) {
        super(props);
        this.searchService = new SearchService(props.context, props.listName);

        this.state = {
            columns: [],
            selectedColumn: "",
            query: "",
            results: [],
            loading: false,
            error: null,
        };
    }

    async componentDidMount() {
        try {
            const columns = await this.searchService.loadColumns();
            this.setState({ columns });
        } catch (error) {
            this.setState({ error: error.message });
        }
    }

    handleSearch = async () => {
        const { selectedColumn, query, columns } = this.state;
        if (!selectedColumn || !query) return;

        this.setState({ loading: true, error: null, results: [] });

        try {
            const selectedColumnInfo = columns.find(col => col.key === selectedColumn);
            if (!selectedColumnInfo) throw new Error("Selected column not found");

            const results = selectedColumnInfo.fieldType === "Lookup"
                ? await this.searchService.handleLookupSearch(selectedColumnInfo, query)
                : await this.searchService.handleStandardSearch(selectedColumn, query);

            this.setState({ results, loading: false });
        } catch (error) {
            this.setState({ error: error.message, loading: false });
        }
    };

    handleColumnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        this.setState({ selectedColumn: option?.key.toString() || "" });
    };

    handleQueryChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        this.setState({ query: newValue || "" });
    };

    render() {
        const { columns, selectedColumn, query, results, loading, error } = this.state;

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
